"""Regression smoke test for the 'user already has HWP open' scenario.

Reproduces the user-reported UX bug where, with a 한/글 window already
visible on screen, calling an MCP edit tool hid the user's document by
applying ``set_visible(False)`` to the shared Hancom process.

The fix (see ``backend/hancom_com.py``) detects a pre-existing Hwp.exe
process and leaves visibility alone in that case. This script verifies
both ends of the behavior:

1. Start a 'user' Hancom instance directly via ``win32com.client.Dispatch``
   with Visible=True (simulating the user opening 한/글 from their menu).
2. Instantiate ``HancomSession`` and run a few MCP-style operations.
3. Assert that the user's Active_XHwpWindow stays Visible=True throughout.
4. Shut down the session WITHOUT killing the user's process (the
   ``_shared_with_user`` guard must protect it).
5. Verify the Hwp.exe is still running afterwards.
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))


def hwp_count() -> int:
    r = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-Process Hwp -ErrorAction SilentlyContinue | "
            "Measure-Object | Select-Object -ExpandProperty Count",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return int((r.stdout or "0").strip() or 0)


def kill_all_hwp() -> None:
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-Process Hwp -ErrorAction SilentlyContinue | "
            "ForEach-Object { Stop-Process -Id $_.Id -Force }",
        ],
        capture_output=True,
    )
    time.sleep(0.5)


def main() -> None:
    # Start from a clean slate so we know exactly which processes belong
    # to the test.
    kill_all_hwp()
    assert hwp_count() == 0, "failed to clean up pre-existing Hwp.exe"

    # --- Simulate user launching 한/글 before the MCP server starts ---
    from hwpx_hwp_mcp.backend import pandas_stub

    pandas_stub.install()
    import win32com.client  # type: ignore[import-not-found]

    user = win32com.client.Dispatch("HWPFrame.HwpObject")
    user.XHwpWindows.Active_XHwpWindow.Visible = True
    user.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
    user_visible_before = bool(user.XHwpWindows.Active_XHwpWindow.Visible)
    assert user_visible_before, "user's window did not start visible"
    count_after_user = hwp_count()
    print(
        f"[setup] user Hancom instance launched, Visible={user_visible_before}, "
        f"Hwp.exe count={count_after_user}"
    )
    assert count_after_user >= 1, "user launch did not spawn Hwp.exe"

    # --- Import HancomSession AFTER user's instance exists ---
    # The module-level singleton ``session`` will be created; its first
    # ``call_sync`` triggers ``_create_on_worker`` which should detect the
    # pre-existing instance and leave visibility alone.
    from hwpx_hwp_mcp.backend.hancom_com import session

    def probe(hwp):
        # Do not touch visibility; just verify we have a working handle.
        return (
            int(hwp.XHwpDocuments.Count),
            bool(hwp.XHwpWindows.Active_XHwpWindow.Visible),
        )

    docs, visible_after_probe = session.call_sync(probe)
    print(
        f"[probe] after MCP dispatch: docs={docs}, Visible={visible_after_probe}"
    )
    assert visible_after_probe is True, (
        "MCP dispatch hid the user's window - set_visible guard failed"
    )
    assert session._shared_with_user is True, (  # noqa: SLF001
        "session did not flag itself as shared_with_user"
    )

    # --- Run an MCP-style operation (add a new doc, insert text) ---
    def edit(hwp):
        hwp.add_doc()
        hwp.insert_text("MCP 작업 중")
        hwp.BreakPara()
        return bool(hwp.XHwpWindows.Active_XHwpWindow.Visible)

    visible_after_edit = session.call_sync(edit)
    print(f"[edit] Visible after add_doc + insert_text: {visible_after_edit}")
    assert visible_after_edit is True, (
        "add_doc / insert_text flipped Visible to False"
    )

    # --- Shut down the session - must NOT kill the shared process ---
    session.shutdown_sync()
    time.sleep(0.5)
    count_after_shutdown = hwp_count()
    print(f"[shutdown] Hwp.exe count after session.shutdown_sync: {count_after_shutdown}")
    assert count_after_shutdown >= 1, (
        "session shutdown killed the user's Hancom process - "
        "_force_kill_tracked_hwp_pid guard failed"
    )

    # --- Cleanup: now we kill the user's instance ourselves ---
    kill_all_hwp()
    assert hwp_count() == 0

    print("OK - shared-instance smoke test passed")


if __name__ == "__main__":
    main()
