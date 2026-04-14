"""Unit tests for backend.errors.translate_com_error."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.errors import (
    HwpArchitectureMismatch,
    HwpError,
    HwpFileLocked,
    HwpNotInstalled,
    translate_com_error,
)


class _FakeComError(Exception):
    """Stand-in for ``pythoncom.com_error``."""


class TestTranslateComError:
    def test_not_installed_from_message(self):
        exc = _FakeComError(-2147221005, "Invalid class string", None, None)
        result = translate_com_error(exc)
        assert isinstance(result, HwpNotInstalled)

    def test_not_installed_from_hwpframe_keyword(self):
        exc = _FakeComError(0, "HWPFrame.HwpObject not found", None, None)
        result = translate_com_error(exc)
        assert isinstance(result, HwpNotInstalled)

    def test_lock_violation_by_hresult(self):
        exc = _FakeComError(0x80030020, "sharing violation", None, None)
        result = translate_com_error(exc)
        assert isinstance(result, HwpFileLocked)

    def test_server_exec_failure_is_architecture_mismatch(self):
        # CO_E_SERVER_EXEC_FAILURE as Windows reports it in signed form.
        exc = _FakeComError(-2146959355, "서버 실행이 실패했습니다.", None, None)
        result = translate_com_error(exc)
        assert isinstance(result, HwpArchitectureMismatch)
        text = str(result)
        assert "32-bit" in text
        assert "python.org" in text.lower()

    def test_generic_fallback(self):
        exc = _FakeComError(0xDEADBEEF, "something else", None, None)
        result = translate_com_error(exc)
        assert isinstance(result, HwpError)
        assert not isinstance(result, (HwpNotInstalled, HwpFileLocked))
        assert "hresult" in str(result).lower()
