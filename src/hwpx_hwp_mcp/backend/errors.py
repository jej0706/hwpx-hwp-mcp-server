"""Custom exception hierarchy and COM error translation.

Tools raise these exceptions; FastMCP converts them into MCP error responses
for the client. Translating HRESULT codes into friendly Korean messages gives
Claude better signal when deciding how to recover.
"""

from __future__ import annotations

from typing import Any


class HwpError(Exception):
    """Base class for all errors surfaced by this MCP server."""


class HwpNotInstalled(HwpError):
    """Raised when ``HWPFrame.HwpObject`` cannot be instantiated."""


class HwpFileLocked(HwpError):
    """Raised when a file is locked by another process (often HWP itself)."""


class HwpUnknownFormat(HwpError):
    """Raised when a requested format is not one we can dispatch to pyhwpx."""


class HwpDocumentNotFound(HwpError):
    """Raised when ``doc_id`` does not match any open document."""


class HwpInvalidPath(HwpError):
    """Raised when a file path fails validation."""


class HwpArchitectureMismatch(HwpError):
    """Raised when 64-bit Python tries to talk to a 32-bit-only Hancom install."""


# HRESULT codes we special-case. ``pythoncom.com_error`` packs these into
# ``args[0]``; the signed/unsigned representation depends on Windows build.
_LOCK_HRESULTS = {
    0x80030020,  # STG_E_LOCKVIOLATION
    0x80030021,  # STG_E_SHAREVIOLATION
    -2147287007,
    -2147287006,
}

# CO_E_SERVER_EXEC_FAILURE — the COM runtime could not launch the LocalServer32
# executable. When this happens with ``HWPFrame.HwpObject`` on a 64-bit Python
# talking to a 32-bit 한/글 install, the real cure is to run under 32-bit Python.
_SERVER_EXEC_FAILURE_HRESULTS = {
    0x80080005,
    -2146959355,
}


def translate_com_error(exc: Exception) -> HwpError:
    """Convert a ``pythoncom.com_error`` (or similar) into a ``HwpError``.

    The input is typed as :class:`Exception` so the module imports cleanly on
    non-Windows hosts used for unit testing.
    """

    args = getattr(exc, "args", ())
    hresult: Any = args[0] if args else None
    msg = ""
    if len(args) >= 2 and isinstance(args[1], str):
        msg = args[1]
    elif len(args) >= 3 and isinstance(args[2], tuple):
        # com_error's excepinfo tuple: (wcode, source, description, ...)
        excepinfo = args[2]
        if len(excepinfo) >= 3 and isinstance(excepinfo[2], str):
            msg = excepinfo[2]

    msg_lower = (msg or "").lower()

    if "hwpframe.hwpobject" in msg_lower or "invalid class string" in msg_lower:
        return HwpNotInstalled(
            "한/글(Hancom HWP)이 설치되어 있지 않거나 COM 등록이 되지 않았습니다. "
            "한/글 설치 후 필요하다면 `regsvr32`로 다시 등록해주세요."
        )

    if isinstance(hresult, int) and (hresult & 0xFFFFFFFF) in {
        h & 0xFFFFFFFF for h in _SERVER_EXEC_FAILURE_HRESULTS
    }:
        return HwpArchitectureMismatch(_architecture_mismatch_message())

    if isinstance(hresult, int) and (hresult & 0xFFFFFFFF) in {
        h & 0xFFFFFFFF for h in _LOCK_HRESULTS
    }:
        return HwpFileLocked(
            "파일이 다른 프로세스에서 열려 있어 접근할 수 없습니다. "
            "한/글에서 해당 파일을 닫은 뒤 다시 시도해주세요."
        )

    hresult_repr = hex(hresult & 0xFFFFFFFF) if isinstance(hresult, int) else repr(hresult)
    return HwpError(
        f"COM 호출 실패: {msg or exc!r} (hresult={hresult_repr})"
    )


def _architecture_mismatch_message() -> str:
    """Build a diagnostic message for CO_E_SERVER_EXEC_FAILURE.

    This covers the by-far most common case: 64-bit Python calling a 32-bit
    Hancom install. We detect Python's bitness at runtime and include a
    concrete remediation path.
    """
    import struct
    import sys

    python_bits = struct.calcsize("P") * 8
    exe = sys.executable

    return (
        f"한/글 COM 서버 기동에 실패했습니다 (CO_E_SERVER_EXEC_FAILURE, 0x80080005).\n"
        f"가장 흔한 원인은 Python 과 한/글의 아키텍처 불일치입니다.\n"
        f"\n"
        f"현재 Python 은 {python_bits}-bit 입니다: {exe}\n"
        f"한/글 2018~2024 는 보통 32-bit (C:\\Program Files (x86)\\Hnc\\...) 로 설치됩니다.\n"
        f"\n"
        f"해결 방법: 32-bit Python 을 설치하고 그 Python 으로 MCP 서버를 실행하세요.\n"
        f"  1) https://www.python.org/downloads/windows/ 에서 'Windows installer (32-bit)' 다운로드\n"
        f"  2) 기존 Python 과 별도 경로로 설치 (예: C:\\Python313-32)\n"
        f"  3) pip install -e D:\\...\\hwpx-hwp-mcp-server\n"
        f"  4) Claude Desktop config 의 command 를 32-bit python.exe 로 지정\n"
        f"\n"
        f"추가로 좀비 Hwp.exe 프로세스가 떠 있다면 작업 관리자에서 정리해주세요."
    )
