"""Session-management tools (category A).

These tools bracket every other call: open a document to obtain a
``doc_id``, save it or export it, then close. ``doc_id`` is the 0-based
index into ``hwp.XHwpDocuments``; callers should refresh their knowledge of
ids after closing a tab, since the collection compacts.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpDocumentNotFound
from ..backend.hancom_com import session
from ..models import (
    CloseResult,
    DocumentRef,
    ListDocumentsResult,
    OpenResult,
    SaveResult,
    to_dict,
)
from ..utils.paths import (
    backup_file,
    ensure_abs_windows_path,
    ensure_existing_file,
    ensure_output_path,
    resolve_save_format,
)


# --------------------------------------------------------------- helpers


def _count(hwp: Any) -> int:
    try:
        return int(hwp.XHwpDocuments.Count)
    except Exception:  # noqa: BLE001
        return 0


def _doc_at(hwp: Any, idx: int) -> Any:
    try:
        # XHwpDocuments is a COM collection; ``Item`` is 0-based in pyhwpx's
        # wrapper conventions. Fallback to Python subscription for safety.
        return hwp.XHwpDocuments.Item(idx)
    except Exception:  # noqa: BLE001
        return hwp.XHwpDocuments[idx]


def _require_doc(hwp: Any, doc_id: int) -> Any:
    count = _count(hwp)
    if doc_id < 0 or doc_id >= count:
        raise HwpDocumentNotFound(
            f"doc_id={doc_id} 가 유효하지 않습니다. 현재 열린 문서 수: {count}"
        )
    switched = hwp.switch_to(doc_id)
    if switched is None:
        raise HwpDocumentNotFound(
            f"doc_id={doc_id} 로 전환할 수 없습니다."
        )
    return switched


def _get_active_doc_id(hwp: Any) -> int:
    """Return the XHwpDocuments index of the currently active document.

    Returns ``-1`` if no document is open.

    Matching logic:
    1. If there is only one document, it is trivially active.
    2. Otherwise compare ``hwp.Path`` (active doc's full path) against each
       ``XHwpDocuments.Item(i).FullName``. First hit wins.
    3. When multiple untitled documents share an empty path, fall back to
       the last index (most recently added is usually active).
    """
    count = _count(hwp)
    if count == 0:
        return -1
    if count == 1:
        return 0

    try:
        active_path = ""
        try:
            active_path = str(hwp.Path or "")
        except Exception:  # noqa: BLE001
            pass

        if active_path:
            for i in range(count):
                try:
                    doc = _doc_at(hwp, i)
                    full = str(getattr(doc, "FullName", "") or "")
                    if full == active_path:
                        return i
                except Exception:  # noqa: BLE001
                    continue
    except Exception:  # noqa: BLE001
        pass

    # Fallback: most-recently-added index (heuristic).
    return count - 1


def _doc_ref_from_active(hwp: Any, doc_id: int) -> DocumentRef:
    title: Optional[str] = None
    path: Optional[str] = None
    page_count: Optional[int] = None
    is_modified: Optional[bool] = None
    try:
        title = str(hwp.Title)
    except Exception:  # noqa: BLE001
        pass
    try:
        path = str(hwp.Path)
    except Exception:  # noqa: BLE001
        pass
    try:
        page_count = int(hwp.PageCount)
    except Exception:  # noqa: BLE001
        pass
    try:
        is_modified = bool(hwp.XHwpDocuments.Item(doc_id).Modified)
    except Exception:  # noqa: BLE001
        pass
    fmt = _format_from_path(path)
    return DocumentRef(
        doc_id=doc_id,
        title=title,
        path=path,
        format=fmt,
        page_count=page_count,
        is_modified=is_modified,
    )


def _format_from_path(path: Optional[str]) -> Optional[str]:
    if not path:
        return None
    ext = Path(path).suffix.lower().lstrip(".")
    return ext.upper() if ext else None


def _find_doc_index_by_path(hwp: Any, target: Path) -> int:
    target_norm = str(target).lower()
    count = _count(hwp)
    for i in range(count):
        try:
            full = str(_doc_at(hwp, i).FullName)
        except Exception:  # noqa: BLE001
            continue
        if full and full.lower() == target_norm:
            return i
    # If we couldn't find by name, assume the most recent doc is the new one.
    return max(0, count - 1)


# --------------------------------------------------------------- tool bodies


def register(mcp: FastMCP) -> None:
    @mcp.tool(
        description=(
            "Open an HWP or HWPX file and return its doc_id. COM-backed: the "
            "file is loaded in the real Hancom HWP engine. Use doc_id for all "
            "subsequent operations; doc_ids reshuffle when tabs close, so "
            "re-run list_open_documents if you are unsure.\n\n"
            "IMPORTANT: Before calling this, check if the user already has "
            "the target file open by calling list_open_documents or "
            "get_active_document first. If the file is already open, reuse "
            "its existing doc_id instead of opening a second copy — opening "
            "an already-open file creates a duplicate window which is "
            "almost never what the user wants."
        ),
    )
    async def open_document(
        file_path: str = Field(..., description="Absolute Windows path to a .hwp or .hwpx file"),
        lock: bool = Field(True, description="Lock the file against external edits while open"),
        read_only: bool = Field(False, description="Open the document in read-only mode"),
    ) -> dict:
        resolved = ensure_existing_file(file_path)

        def _do(hwp: Any) -> OpenResult:
            arg_parts: list[str] = []
            if lock:
                arg_parts.append("lock:true")
            if read_only:
                arg_parts.append("readonly:true")
            arg = ";".join(arg_parts)
            if not hwp.open(str(resolved), format="", arg=arg):
                raise RuntimeError(f"hwp.open returned False for {resolved}")
            doc_id = _find_doc_index_by_path(hwp, resolved)
            _require_doc(hwp, doc_id)
            return OpenResult(**_doc_ref_from_active(hwp, doc_id).model_dump())

        result = await session.call(_do)
        return to_dict(result)

    @mcp.tool(
        description=(
            "Get a workable document to edit in 한글/HWP. "
            "⚠️ DEFAULT BEHAVIOR: returns the currently-active document if one "
            "is already open, so the user's existing work is reused and no "
            "extra window is spawned. THIS IS ALMOST ALWAYS WHAT YOU WANT.\n\n"
            "Only pass `force_new=True` if the user EXPLICITLY asks for a "
            "brand-new blank document that must coexist with their existing "
            "work (e.g. 'create a second empty document in addition to the "
            "one I have open'). Do NOT pass force_new=True just because the "
            "user said '새 문서 만들고 표 넣어줘' — in that phrasing '새 "
            "문서' usually refers to the blank one already on screen.\n\n"
            "If you are unsure whether a document is already open, call "
            "`get_active_document` or `list_open_documents` first — they "
            "have no side effects and will tell you the current state.\n\n"
            "When a new document actually has to be created, it is added as "
            "a NEW TAB inside the existing 한글 window (not a separate "
            "window) to keep the workspace uncluttered."
        ),
    )
    async def create_new_document(
        force_new: bool = Field(
            False,
            description=(
                "Default False — reuse the active document if any exists. "
                "Set to True ONLY when the user explicitly wants an "
                "additional brand-new blank document alongside whatever is "
                "already open. Most '한글에 ...써줘' / '표 만들어줘' "
                "requests should leave this False."
            ),
        ),
    ) -> dict:
        def _do(hwp: Any) -> OpenResult:
            # Default path: reuse the active document if any exists.
            if not force_new:
                active_idx = _get_active_doc_id(hwp)
                if active_idx >= 0:
                    _require_doc(hwp, active_idx)
                    ref = _doc_ref_from_active(hwp, active_idx).model_dump()
                    return OpenResult(**ref)

            # Force-new path: add a new tab inside the existing window.
            # (We always use add_tab, never add_doc, because a separate
            # window is even more disruptive to the user's workspace.)
            hwp.add_tab()
            doc_id = max(0, _count(hwp) - 1)
            _require_doc(hwp, doc_id)
            return OpenResult(**_doc_ref_from_active(hwp, doc_id).model_dump())

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Return the currently active document's doc_id and metadata "
            "without creating or opening anything. Call this FIRST in any "
            "editing workflow — if the user already has 한글 open, you "
            "should reuse the active document rather than creating a new "
            "one (which would spawn an extra HWP window/tab).\n\n"
            "Raises HwpDocumentNotFound if no document is currently open; "
            "in that case fall back to create_new_document or open_document."
        ),
    )
    async def get_active_document() -> dict:
        def _do(hwp: Any) -> DocumentRef:
            active_idx = _get_active_doc_id(hwp)
            if active_idx < 0:
                raise HwpDocumentNotFound(
                    "현재 열려 있는 한글 문서가 없습니다. "
                    "create_new_document 또는 open_document 를 먼저 호출하세요."
                )
            _require_doc(hwp, active_idx)
            return _doc_ref_from_active(hwp, active_idx)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Save the document at its current path. If backup=True (default) a "
            ".bak copy of the existing file is made first. Use save_as to pick "
            "a new path or change format."
        ),
    )
    async def save_document(
        doc_id: int = Field(..., description="Document index from open_document"),
        backup: bool = Field(True, description="Create a .bak copy of the existing file"),
    ) -> dict:
        def _do(hwp: Any) -> SaveResult:
            _require_doc(hwp, doc_id)
            current_path = ""
            try:
                current_path = str(hwp.Path)
            except Exception:  # noqa: BLE001
                pass

            backup_path: Optional[str] = None
            if backup and current_path:
                try:
                    made = backup_file(Path(current_path))
                    backup_path = str(made) if made else None
                except Exception:  # noqa: BLE001
                    backup_path = None

            ok = bool(hwp.save(save_if_dirty=True))
            return SaveResult(
                saved=ok,
                path=current_path,
                format=_format_from_path(current_path),
                backup_path=backup_path,
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Save the active document to output_path, auto-detecting the "
            "target format from the extension (.hwp/.hwpx/.pdf/.html/.docx) "
            "or from an explicit format argument. Use this to convert HWP ↔ "
            "HWPX ↔ PDF ↔ DOCX."
        ),
    )
    async def save_as(
        doc_id: int = Field(..., description="Document index from open_document"),
        output_path: str = Field(..., description="Absolute destination path"),
        format: str = Field(
            "auto",
            description="auto | HWP | HWPX | PDF | HTML | DOCX (default: auto by extension)",
        ),
        create_dirs: bool = Field(
            False, description="Create the parent directory if it does not exist"
        ),
    ) -> dict:
        out = ensure_output_path(output_path, create_dirs=create_dirs)
        fmt = resolve_save_format(format, out)

        def _do(hwp: Any) -> SaveResult:
            _require_doc(hwp, doc_id)
            # When fmt == "" pyhwpx defaults to 'HWP' so we always pass a
            # non-empty string; the extension overrides for hwpx/pdf anyway.
            arg = ""
            ok = bool(hwp.save_as(str(out), format=fmt or "HWP", arg=arg))
            return SaveResult(
                saved=ok,
                path=str(out),
                format=fmt or _format_from_path(str(out)),
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Close a document by doc_id. Set save=True to save before closing. "
            "After closing, remaining doc_ids shift down; re-fetch them with "
            "list_open_documents."
        ),
    )
    async def close_document(
        doc_id: int = Field(..., description="Document index to close"),
        save: bool = Field(False, description="If True, save before closing"),
    ) -> dict:
        def _do(hwp: Any) -> CloseResult:
            _require_doc(hwp, doc_id)
            doc = _doc_at(hwp, doc_id)
            try:
                doc.Close(isDirty=bool(save))
            except Exception:  # noqa: BLE001
                # Older wrappers want no keyword args.
                doc.Close(bool(save))
            return CloseResult(closed=True, doc_id=doc_id)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "List every document currently open in the Hancom session, with "
            "its doc_id, title, path, and modified flag."
        ),
    )
    async def list_open_documents() -> dict:
        def _do(hwp: Any) -> ListDocumentsResult:
            count = _count(hwp)
            out: list[DocumentRef] = []
            for i in range(count):
                title: Optional[str] = None
                path: Optional[str] = None
                modified: Optional[bool] = None
                try:
                    doc = _doc_at(hwp, i)
                    path = str(doc.FullName) if doc.FullName else None
                    modified = bool(doc.Modified)
                except Exception:  # noqa: BLE001
                    pass
                # Title is only available on the active doc.
                if i == 0:
                    try:
                        title = str(hwp.Title)
                    except Exception:  # noqa: BLE001
                        pass
                out.append(
                    DocumentRef(
                        doc_id=i,
                        title=title,
                        path=path,
                        format=_format_from_path(path),
                        is_modified=modified,
                    )
                )
            return ListDocumentsResult(documents=out)

        return to_dict(await session.call(_do))
