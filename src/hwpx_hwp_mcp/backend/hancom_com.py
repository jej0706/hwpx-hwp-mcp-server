"""Hancom HWP COM session.

Design notes
------------

Hancom HWP's COM object lives in a **Single-Threaded Apartment (STA)**. All
calls into ``HWPFrame.HwpObject`` must originate from the same thread that
invoked ``pythoncom.CoInitialize``. ``asyncio.to_thread`` uses a shared
thread-pool whose worker thread can change between calls, which would crash
the COM client. We therefore pin a dedicated ``ThreadPoolExecutor`` with
``max_workers=1`` and route every pyhwpx call through it.

The session is a module-level singleton so that the long-lived stdio MCP
process reuses a single Hancom instance (boot time is several seconds).
``_ensure`` performs a cheap health-check before handing the instance to a
caller; if the instance looks dead, it is recreated on the same worker
thread.
"""

from __future__ import annotations

import asyncio
import atexit
import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, Optional, TypeVar

# Install the pandas stub BEFORE anything imports pyhwpx. pyhwpx's core.py
# does ``import pandas as pd`` at module load; on 32-bit Python 3.10+ there
# are no prebuilt pandas wheels, so we satisfy the import with a lightweight
# sentinel module. Real pandas wins when present.
from . import pandas_stub as _pandas_stub

_pandas_stub.install()

from .errors import HwpError, HwpNotInstalled, translate_com_error  # noqa: E402

logger = logging.getLogger(__name__)

T = TypeVar("T")


class HancomSession:
    """Owns a single Hancom ``Hwp`` instance and serializes access to it."""

    def __init__(self) -> None:
        self._executor = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="hwp-com"
        )
        self._hwp: Optional[Any] = None
        self._init_lock = threading.Lock()
        self._initialized_com = False
        self._shutdown = False
        self._tracked_pid: Optional[int] = None
        self._shared_with_user: bool = False

    # ------------------------------------------------------------------ internal

    def _create_on_worker(self) -> Any:
        """Create a fresh ``Hwp`` instance. Called on the executor thread only.

        IMPORTANT: Hancom's ``HwpObject`` COM server is registered as
        ``MultipleUse`` LocalServer32, which means all Python clients connect
        to the *same* ``Hwp.exe`` process. That process has a single
        ``Active_XHwpWindow`` shared between us and any 한/글 UI the user
        already has open. Calling ``set_visible(False)`` under those
        conditions hides the user's own document.

        We therefore detect whether 한/글 was already running before we
        dispatched. If it was, we leave visibility alone (the user owns that
        window). If not, we assume the instance is ours alone and can safely
        hide it so no stray 한/글 window pops up on the user's desktop.
        """
        import pythoncom  # type: ignore[import-not-found]
        from pyhwpx import Hwp  # type: ignore[import-not-found]

        if not self._initialized_com:
            pythoncom.CoInitialize()
            self._initialized_com = True

        hancom_was_already_running = _hwp_process_count() > 0

        try:
            hwp = Hwp(new=False, visible=True, register_module=True)
        except Exception as exc:  # noqa: BLE001
            raise translate_com_error(exc) from exc

        if not hancom_was_already_running:
            try:
                hwp.set_visible(False)
            except Exception:  # noqa: BLE001 - visibility tweak is best-effort
                logger.debug("set_visible(False) failed; continuing")
        else:
            logger.info(
                "Detected pre-existing Hancom HWP instance; leaving window "
                "visibility alone to avoid hiding user documents."
            )
        self._shared_with_user = hancom_was_already_running

        try:
            # MESSAGE_BOX_MODE = 0x00000010 (silent / suppress all dialog boxes)
            hwp.set_message_box_mode(0x10)
        except Exception:  # noqa: BLE001
            logger.debug("set_message_box_mode(0x10) failed; continuing")

        # Track the Hwp.exe PID so ``_force_kill_tracked_hwp_pid`` can nuke it
        # if graceful quit fails. Best-effort fall-back; because Hancom is
        # MultipleUse, the PID may belong to a process the user started, so
        # we only taskkill it as a last resort on shutdown AND only when the
        # instance was ours (``shared_with_user`` is False).
        try:
            self._tracked_pid = _find_latest_hwp_pid()
        except Exception:  # noqa: BLE001
            self._tracked_pid = None

        return hwp

    def _ensure_on_worker(self) -> Any:
        """Return a live ``Hwp``; recreate it if the previous one is dead."""
        if self._shutdown:
            raise HwpError("Hancom session has been shut down.")

        if self._hwp is None:
            self._hwp = self._create_on_worker()
            return self._hwp

        try:
            # Cheap probe: touching ``XHwpDocuments.Count`` fails if the
            # Hancom process went away.
            _ = self._hwp.XHwpDocuments.Count
        except Exception:  # noqa: BLE001
            logger.warning("Hancom instance looked dead; recreating")
            self._hwp = None
            self._hwp = self._create_on_worker()
        return self._hwp

    def _run(self, fn: Callable[[Any], T]) -> T:
        """Body executed inside the dedicated COM worker thread."""
        try:
            hwp = self._ensure_on_worker()
            return fn(hwp)
        except HwpError:
            raise
        except Exception as exc:
            # Any raw COM/win32 error becomes a HwpError. We also reset the
            # cached instance so the next call gets a fresh one.
            self._hwp = None
            raise translate_com_error(exc) from exc

    # ------------------------------------------------------------------ public

    async def call(self, fn: Callable[[Any], T]) -> T:
        """Run ``fn(hwp)`` on the COM worker thread and await the result."""
        if self._shutdown:
            raise HwpError("Hancom session has been shut down.")
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(self._executor, self._run, fn)

    def call_sync(self, fn: Callable[[Any], T]) -> T:
        """Synchronous variant used by smoke scripts and tests."""
        if self._shutdown:
            raise HwpError("Hancom session has been shut down.")
        future = self._executor.submit(self._run, fn)
        return future.result()

    async def shutdown(self) -> None:
        """Quit the Hancom process and release COM resources."""
        if self._shutdown:
            return
        self._shutdown = True
        loop = asyncio.get_running_loop()
        try:
            await loop.run_in_executor(self._executor, self._shutdown_on_worker)
        except Exception:  # noqa: BLE001
            logger.exception("Error during Hancom shutdown; continuing")
        finally:
            self._executor.shutdown(wait=False, cancel_futures=True)

    def shutdown_sync(self, *, timeout: float = 10.0) -> None:
        """Synchronous shutdown for smoke tests and ``atexit``.

        Routes ``_shutdown_on_worker`` through the executor so the STA
        thread ownership is respected. Safe to call multiple times.
        """
        if self._shutdown:
            return
        self._shutdown = True
        try:
            future = self._executor.submit(self._shutdown_on_worker)
            future.result(timeout=timeout)
        except Exception:  # noqa: BLE001
            pass
        try:
            self._executor.shutdown(wait=False, cancel_futures=True)
        except Exception:  # noqa: BLE001
            pass

    def _shutdown_on_worker(self) -> None:
        if self._hwp is not None:
            if self._shared_with_user:
                # User owns this Hancom process — don't quit it and don't
                # close their documents. Just drop our reference and let
                # COM's refcounting settle.
                logger.info(
                    "Skipping hwp.quit() because Hancom is shared with user."
                )
            else:
                try:
                    self._hwp.quit(save=False)
                except Exception:  # noqa: BLE001
                    logger.debug("hwp.quit failed; continuing")
            self._hwp = None

        if self._initialized_com:
            try:
                import pythoncom  # type: ignore[import-not-found]

                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                logger.debug("CoUninitialize failed; continuing")
            finally:
                self._initialized_com = False


def _find_latest_hwp_pid() -> Optional[int]:
    """Return the PID of the most recently started Hwp.exe, if any.

    Used as a tracking heuristic so ``_force_kill_tracked_hwp_pid`` has a
    concrete target when graceful shutdown fails. On Windows, queries via
    PowerShell; returns None on any error.
    """
    try:
        import subprocess

        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "Get-Process Hwp -ErrorAction SilentlyContinue | "
                "Sort-Object StartTime -Descending | "
                "Select-Object -First 1 -ExpandProperty Id",
            ],
            capture_output=True,
            text=True,
            timeout=5,
        )
        pid_str = (result.stdout or "").strip()
        return int(pid_str) if pid_str else None
    except Exception:  # noqa: BLE001
        return None


def _hwp_process_count() -> int:
    """Return the number of running ``Hwp.exe`` processes on the system.

    Used to decide whether the Hancom instance we're about to talk to is
    already owned by the user (visible) or was started by us (hideable).
    Returns 0 on any error so we default to "probably ours" behavior.
    """
    try:
        import subprocess

        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "Get-Process Hwp -ErrorAction SilentlyContinue | "
                "Measure-Object | Select-Object -ExpandProperty Count",
            ],
            capture_output=True,
            text=True,
            timeout=5,
        )
        return int((result.stdout or "0").strip() or 0)
    except Exception:  # noqa: BLE001
        return 0


# Module-level singleton used by every tool module.
session = HancomSession()


def _atexit_shutdown() -> None:
    """Last-chance cleanup in case FastMCP lifespan did not run.

    Python's ``concurrent.futures`` cleanup runs before module-level atexit
    handlers, which means ``session._executor.submit`` may silently drop our
    job. We still try the graceful path first, and if the executor is dead
    we fall back to taskkill on the tracked Hancom PID — ugly but prevents
    zombie Hwp.exe processes when users Ctrl-C out of the stdio server.
    """
    session.shutdown_sync(timeout=5)
    _force_kill_tracked_hwp_pid()


def _force_kill_tracked_hwp_pid() -> None:
    """Last-resort taskkill for our Hwp.exe, but NEVER when shared with user.

    If ``HancomSession._shared_with_user`` is True, the Hwp.exe process
    was already running when our MCP server started — that almost
    certainly means the user launched it for their own work. Killing it
    would destroy unsaved documents. Bail out silently in that case and
    let the user close their 한/글 manually if needed.
    """
    if getattr(session, "_shared_with_user", False):
        return
    pid = getattr(session, "_tracked_pid", None)
    if not pid:
        return
    try:
        import subprocess

        subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            capture_output=True,
            timeout=5,
        )
    except Exception:  # noqa: BLE001
        pass


atexit.register(_atexit_shutdown)


__all__ = ["HancomSession", "session", "HwpError", "HwpNotInstalled"]
