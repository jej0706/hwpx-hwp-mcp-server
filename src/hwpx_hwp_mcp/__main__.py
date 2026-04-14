"""Entry point: ``python -m hwpx_hwp_mcp``.

Runs preflight checks (Windows + Hancom HWP COM registration) before handing
control to the FastMCP stdio server.
"""

from __future__ import annotations

import struct
import sys


def _preflight() -> None:
    if sys.platform != "win32":
        sys.stderr.write(
            "hwpx-hwp-mcp requires Windows with Hancom HWP installed.\n"
            f"Current platform: {sys.platform}\n"
        )
        raise SystemExit(2)

    python_bits = struct.calcsize("P") * 8

    # Early architecture-mismatch warning. 한/글 is almost always 32-bit.
    if python_bits == 64:
        if _hancom_looks_32bit_only():
            sys.stderr.write(
                "WARNING: 64-bit Python detected, but Hancom HWP appears to be "
                "installed as 32-bit only.\n"
                "  Python:     64-bit ({exe})\n"
                "  한/글:      32-bit (CLSID registered only under Wow6432Node)\n"
                "\n"
                "COM instantiation will likely fail with CO_E_SERVER_EXEC_FAILURE.\n"
                "Install 32-bit Python from https://www.python.org/downloads/windows/\n"
                "and run this MCP server under that interpreter.\n"
                "Continuing anyway — the real error will follow.\n".format(exe=sys.executable)
            )

    # Light COM probe: fail fast with a friendly message if HWP is not installed.
    try:
        import win32com.client  # type: ignore[import-not-found]

        win32com.client.Dispatch("HWPFrame.HwpObject")
    except Exception as exc:  # noqa: BLE001 - we want to surface any COM failure
        from .backend.errors import HwpArchitectureMismatch, translate_com_error

        translated = translate_com_error(exc)
        if isinstance(translated, HwpArchitectureMismatch):
            sys.stderr.write(str(translated) + "\n")
            raise SystemExit(4) from exc

        sys.stderr.write(
            "Failed to instantiate HWPFrame.HwpObject COM object.\n"
            "Install Hancom Office (한/글) and verify COM registration.\n"
            f"Underlying error: {exc}\n"
        )
        raise SystemExit(3) from exc


def _hancom_looks_32bit_only() -> bool:
    """Best-effort check: is the Hancom CLSID registered only in Wow6432Node?

    Returns False on any error so we never block startup on a heuristic.
    """
    try:
        import winreg  # type: ignore[import-not-found]
    except Exception:  # noqa: BLE001
        return False

    clsid = "{2291CF00-64A1-4877-A9B4-68CFE89612D6}"
    native_path = rf"SOFTWARE\Classes\CLSID\{clsid}\LocalServer32"
    wow_path = rf"SOFTWARE\Classes\Wow6432Node\CLSID\{clsid}\LocalServer32"

    has_native = _registry_key_exists(winreg.HKEY_LOCAL_MACHINE, native_path)
    has_wow = _registry_key_exists(winreg.HKEY_LOCAL_MACHINE, wow_path)
    return (not has_native) and has_wow


def _registry_key_exists(hive: int, subkey: str) -> bool:
    try:
        import winreg  # type: ignore[import-not-found]

        with winreg.OpenKey(hive, subkey):
            return True
    except OSError:
        return False
    except Exception:  # noqa: BLE001
        return False


def main() -> None:
    _preflight()
    # Import after the preflight so that an import-time side effect of
    # pyhwpx does not mask the friendly diagnostic above.
    from .server import mcp

    mcp.run()


if __name__ == "__main__":
    main()
