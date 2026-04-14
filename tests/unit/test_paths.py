"""Unit tests for utils.paths — pure logic, no COM required."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.errors import HwpInvalidPath, HwpUnknownFormat
from hwpx_hwp_mcp.utils.paths import (
    backup_file,
    ensure_abs_windows_path,
    resolve_save_format,
)


class TestEnsureAbsWindowsPath:
    def test_accepts_drive_letter_backslash(self):
        result = ensure_abs_windows_path("C:\\Users\\me\\doc.hwpx")
        assert str(result).lower().startswith("c:")

    def test_accepts_forward_slashes(self):
        result = ensure_abs_windows_path("C:/Users/me/doc.hwpx")
        assert "Users" in str(result)

    def test_rejects_relative(self):
        with pytest.raises(HwpInvalidPath):
            ensure_abs_windows_path("docs\\file.hwp")

    def test_rejects_empty(self):
        with pytest.raises(HwpInvalidPath):
            ensure_abs_windows_path("")

    def test_rejects_non_string(self):
        with pytest.raises(HwpInvalidPath):
            ensure_abs_windows_path(None)  # type: ignore[arg-type]


class TestResolveSaveFormat:
    def test_auto_hwp_from_extension(self, tmp_path):
        assert resolve_save_format("auto", tmp_path / "x.hwp") == "HWP"

    def test_auto_hwpx_from_extension(self, tmp_path):
        assert resolve_save_format("auto", tmp_path / "x.hwpx") == "HWPX"

    def test_auto_pdf(self, tmp_path):
        assert resolve_save_format("auto", tmp_path / "x.pdf") == "PDF"

    def test_auto_docx(self, tmp_path):
        assert resolve_save_format("auto", tmp_path / "x.docx") == "OOXML"

    def test_auto_unknown_extension_defaults_hwp(self, tmp_path):
        assert resolve_save_format("auto", tmp_path / "x.bin") == "HWP"

    def test_explicit_override(self, tmp_path):
        # Explicit beats extension.
        assert resolve_save_format("HWPX", tmp_path / "x.hwp") == "HWPX"

    def test_unknown_format_raises(self, tmp_path):
        with pytest.raises(HwpUnknownFormat):
            resolve_save_format("excel", tmp_path / "x.xlsx")


class TestBackupFile:
    def test_creates_bak_copy(self, tmp_path):
        original = tmp_path / "doc.hwpx"
        original.write_bytes(b"payload")
        backup = backup_file(original)
        assert backup is not None
        assert backup.exists()
        assert backup.read_bytes() == b"payload"
        assert backup.name == "doc.hwpx.bak"

    def test_missing_file_returns_none(self, tmp_path):
        assert backup_file(tmp_path / "nope.hwp") is None

    def test_timestamped_variant(self, tmp_path):
        original = tmp_path / "doc.hwp"
        original.write_bytes(b"x")
        backup = backup_file(original, timestamped=True)
        assert backup is not None
        assert backup.name.startswith("doc.hwp.bak.")
