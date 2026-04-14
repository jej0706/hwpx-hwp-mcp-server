"""Windows path validation and backup helpers.

Tools must never pass relative paths to pyhwpx — ``insert_picture`` and
``open`` silently resolve them against a Hancom-internal working directory
which surprises callers. We force absolute Windows paths here and do the
``Path.exists`` / ``parent.mkdir`` plumbing once.
"""

from __future__ import annotations

import datetime as _dt
import shutil
from pathlib import Path, PureWindowsPath
from typing import Iterable, Optional

from ..backend.errors import HwpInvalidPath, HwpUnknownFormat

# Formats accepted by ``pyhwpx.Hwp.save_as``. Keys are the user-facing
# lowercase names; values are the ``format`` argument expected by pyhwpx.
# ``save_as`` also inspects the file extension for HWPX/PDF, so passing the
# right extension already dispatches correctly — these are only used when the
# caller forces a particular format.
SAVE_FORMATS: dict[str, str] = {
    "auto": "",
    "hwp": "HWP",
    "hwpx": "HWPX",
    "pdf": "PDF",
    "html": "HTML",
    "docx": "OOXML",
    "text": "TEXT",
    "unicode": "UNICODE",
}

# Extensions we know how to save via the normal `save_as` path.
KNOWN_EXTENSIONS = {".hwp", ".hwpx", ".pdf", ".html", ".htm", ".docx", ".txt"}


def ensure_abs_windows_path(raw: str) -> Path:
    """Return an absolute Windows ``Path`` or raise ``HwpInvalidPath``."""
    if not raw or not isinstance(raw, str):
        raise HwpInvalidPath("파일 경로는 비어 있지 않은 문자열이어야 합니다.")

    # Normalize both slash styles.
    candidate = PureWindowsPath(raw.replace("/", "\\"))
    if not candidate.is_absolute():
        raise HwpInvalidPath(
            f"절대 경로(드라이브 문자 포함)가 필요합니다. 받은 값: {raw!r}"
        )

    # ``Path(str(candidate))`` keeps Windows semantics on Windows and still
    # works on non-Windows hosts used for unit tests.
    return Path(str(candidate))


def ensure_existing_file(raw: str) -> Path:
    path = ensure_abs_windows_path(raw)
    if not path.exists():
        raise HwpInvalidPath(f"파일이 존재하지 않습니다: {path}")
    if not path.is_file():
        raise HwpInvalidPath(f"파일이 아닙니다: {path}")
    return path


def ensure_output_path(raw: str, *, create_dirs: bool = False) -> Path:
    path = ensure_abs_windows_path(raw)
    parent = path.parent
    if not parent.exists():
        if create_dirs:
            parent.mkdir(parents=True, exist_ok=True)
        else:
            raise HwpInvalidPath(
                f"출력 디렉토리가 존재하지 않습니다: {parent} "
                "(create_dirs=True 옵션으로 자동 생성 가능)"
            )
    return path


def backup_file(path: Path, *, timestamped: bool = False) -> Optional[Path]:
    """Copy ``path`` to ``<path>.bak`` (or a timestamped suffix)."""
    if not path.exists():
        return None
    if timestamped:
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = path.with_suffix(path.suffix + f".bak.{stamp}")
    else:
        backup = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, backup)
    return backup


def resolve_save_format(format_name: str, path: Path) -> str:
    """Return the ``format`` string expected by ``pyhwpx.Hwp.save_as``.

    ``pyhwpx`` already dispatches on the file extension for HWPX and PDF, so
    the empty string is safe for the "auto" case — we just pass it through.
    """
    key = (format_name or "auto").lower()
    if key not in SAVE_FORMATS:
        raise HwpUnknownFormat(
            f"지원하지 않는 포맷입니다: {format_name!r}. "
            f"허용 값: {sorted(SAVE_FORMATS)}"
        )
    if key == "auto":
        ext = path.suffix.lower()
        if ext in (".hwp",):
            return "HWP"
        if ext in (".hwpx",):
            return "HWPX"
        if ext in (".pdf",):
            return "PDF"
        if ext in (".html", ".htm"):
            return "HTML"
        if ext in (".docx",):
            return "OOXML"
        if ext in (".txt",):
            return "TEXT"
        # Unknown extension: let pyhwpx default to HWP.
        return "HWP"
    return SAVE_FORMATS[key]


def iter_input_files(
    paths: Optional[Iterable[str]] = None,
    *,
    folder: Optional[str] = None,
    glob: str = "*.hwp*",
) -> list[Path]:
    """Resolve a list of input paths for batch tools."""
    resolved: list[Path] = []
    if paths:
        for raw in paths:
            resolved.append(ensure_existing_file(raw))
    if folder:
        root = ensure_abs_windows_path(folder)
        if not root.exists() or not root.is_dir():
            raise HwpInvalidPath(f"폴더가 존재하지 않습니다: {root}")
        resolved.extend(sorted(p for p in root.glob(glob) if p.is_file()))
    if not resolved:
        raise HwpInvalidPath(
            "처리할 파일을 찾지 못했습니다. input_paths 또는 folder/glob 을 지정해주세요."
        )
    # Deduplicate while preserving order.
    seen: set[str] = set()
    unique: list[Path] = []
    for p in resolved:
        key = str(p).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(p)
    return unique
