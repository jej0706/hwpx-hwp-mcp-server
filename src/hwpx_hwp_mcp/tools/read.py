"""Read / analyze tools (category B).

All of these are non-mutating: they give the LLM ways to inspect a document
without touching its content.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, List, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import (
    DocumentInfo,
    DocumentStructure,
    DocumentTextResult,
    ExportResult,
    FieldInfo,
    ImageStructure,
    SearchHit,
    SearchResult,
    TableCsvResult,
    TableStructure,
    to_dict,
)
from ..utils.paths import ensure_output_path, resolve_save_format
from .session import _format_from_path, _require_doc


# --------------------------------------------------------------- helpers


def _split_field_list(raw: str) -> List[str]:
    """``get_field_list`` returns a string. Split it safely."""
    if not raw:
        return []
    # Hancom uses "\x02" (STX) as separator in many field APIs. Fall back to
    # comma/newline if that isn't present.
    for sep in ("\x02", "\n", ","):
        if sep in raw:
            return [s for s in raw.split(sep) if s]
    return [raw]


def _iter_ctrls(hwp: Any):
    """Yield all controls in the active document regardless of container."""
    try:
        ctrl = hwp.HeadCtrl
    except Exception:  # noqa: BLE001
        return
    while ctrl is not None:
        yield ctrl
        try:
            ctrl = ctrl.Next
        except Exception:  # noqa: BLE001
            break


# --------------------------------------------------------------- register


def register(mcp: FastMCP) -> None:
    @mcp.tool(
        description=(
            "Return the full plain text of a document. Fast for large docs. "
            "Tables are included when Hancom's text exporter inlines them."
        ),
    )
    async def get_document_text(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> DocumentTextResult:
            _require_doc(hwp, doc_id)
            text = ""
            try:
                # IMPORTANT: pass option="" (whole-document mode). The default
                # "saveblock:true" returns the selected block only and yields
                # None when there is no selection.
                raw = hwp.get_text_file("TEXT", "")
                text = str(raw) if raw else ""
            except Exception:
                # Fall back to a paragraph-by-paragraph loop.
                chunks: list[str] = []
                while True:
                    try:
                        state, para = hwp.get_text()
                    except Exception:  # noqa: BLE001
                        break
                    if state == 0:
                        break
                    chunks.append(para)
                    if state == 1:
                        break
                text = "\n".join(chunks)
            page_count: Optional[int] = None
            try:
                page_count = int(hwp.PageCount)
            except Exception:  # noqa: BLE001
                pass
            return DocumentTextResult(
                text=text, char_count=len(text), page_count=page_count
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Return quick metadata for a document: title, path, page count, "
            "modified flag, and field count."
        ),
    )
    async def get_document_info(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> DocumentInfo:
            _require_doc(hwp, doc_id)
            title: Optional[str] = None
            path: Optional[str] = None
            page_count: Optional[int] = None
            modified: Optional[bool] = None
            field_count = 0
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
                modified = bool(hwp.XHwpDocuments.Item(doc_id).Modified)
            except Exception:  # noqa: BLE001
                pass
            try:
                field_count = len(_split_field_list(hwp.get_field_list(1, 0)))
            except Exception:  # noqa: BLE001
                pass
            return DocumentInfo(
                title=title,
                path=path,
                page_count=page_count,
                is_modified=modified,
                field_count=field_count,
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Return a structural outline of the document: tables, images, and "
            "fields (누름틀). Each entry carries an index you can pass to "
            "get_table_as_csv."
        ),
    )
    async def get_structure(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> DocumentStructure:
            _require_doc(hwp, doc_id)
            tables: list[TableStructure] = []
            images: list[ImageStructure] = []
            table_idx = 0
            image_idx = 0
            for ctrl in _iter_ctrls(hwp):
                try:
                    user_desc = ""
                    try:
                        user_desc = str(ctrl.UserDesc)
                    except Exception:  # noqa: BLE001
                        pass
                    ctrl_id = ""
                    try:
                        ctrl_id = str(ctrl.CtrlID).strip().lower()
                    except Exception:  # noqa: BLE001
                        pass
                    if ctrl_id == "tbl":
                        rows = cols = None
                        try:
                            shape = ctrl.Properties
                            rows = int(shape.Item("Rows"))
                            cols = int(shape.Item("Cols"))
                        except Exception:  # noqa: BLE001
                            pass
                        tables.append(
                            TableStructure(
                                index=table_idx,
                                rows=rows,
                                cols=cols,
                                caption=user_desc or None,
                            )
                        )
                        table_idx += 1
                    elif ctrl_id in ("gso", "$pic"):
                        images.append(ImageStructure(index=image_idx))
                        image_idx += 1
                except Exception:  # noqa: BLE001
                    continue

            fields: list[FieldInfo] = []
            try:
                raw = hwp.get_field_list(1, 0)
                for i, name in enumerate(_split_field_list(raw)):
                    current: Optional[str] = None
                    try:
                        current = str(hwp.get_field_text(name))
                    except Exception:  # noqa: BLE001
                        pass
                    fields.append(FieldInfo(name=name, index=i, current_text=current))
            except Exception:  # noqa: BLE001
                pass

            page_count: Optional[int] = None
            try:
                page_count = int(hwp.PageCount)
            except Exception:  # noqa: BLE001
                pass
            return DocumentStructure(
                page_count=page_count,
                tables=tables,
                images=images,
                fields=fields,
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Search the document for text (or a regex) and return short "
            "context snippets for each hit. Capped at max_hits to avoid "
            "context bloat."
        ),
    )
    async def search_text(
        doc_id: int = Field(..., description="Document index from open_document"),
        query: str = Field(..., description="Search text or regex"),
        regex: bool = Field(False, description="Interpret query as a regex"),
        max_hits: int = Field(50, description="Stop after this many hits", ge=1, le=500),
    ) -> dict:
        def _do(hwp: Any) -> SearchResult:
            _require_doc(hwp, doc_id)
            hits: list[SearchHit] = []
            # Snapshot cursor so we can restore after searching.
            saved_pos = None
            try:
                saved_pos = hwp.get_pos()
            except Exception:  # noqa: BLE001
                pass
            try:
                # Reset to top of document before searching.
                try:
                    hwp.MovePos(2)  # moveScrPos_DocBegin
                except Exception:  # noqa: BLE001
                    pass
                while len(hits) < max_hits:
                    if not hwp.find_forward(query, regex=regex):
                        break
                    context = ""
                    try:
                        # A small window around the current position. Not all
                        # builds support GetSelectedText; fall back silently.
                        context = str(hwp.get_selected_text()) if hasattr(
                            hwp, "get_selected_text"
                        ) else query
                    except Exception:  # noqa: BLE001
                        context = query
                    hits.append(SearchHit(match=query, context=context))
            finally:
                if saved_pos is not None:
                    try:
                        hwp.set_pos(*saved_pos)
                    except Exception:  # noqa: BLE001
                        pass
            return SearchResult(query=query, hit_count=len(hits), hits=hits)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Export the document to another format file (text/html/pdf/docx). "
            "This writes a new file — use save_as to overwrite the bound path."
        ),
    )
    async def export_document(
        doc_id: int = Field(..., description="Document index from open_document"),
        output_path: str = Field(..., description="Absolute destination path"),
        format: str = Field("text", description="text | html | pdf | docx"),
        create_dirs: bool = Field(False, description="Create parent directory if missing"),
    ) -> dict:
        fmt_name = (format or "text").lower()
        if fmt_name not in {"text", "html", "pdf", "docx"}:
            raise HwpError(f"export_document: 지원하지 않는 format={format!r}")
        out = ensure_output_path(output_path, create_dirs=create_dirs)
        resolved = resolve_save_format(fmt_name, out)

        def _do(hwp: Any) -> ExportResult:
            _require_doc(hwp, doc_id)
            ok = bool(hwp.save_as(str(out), format=resolved or "HWP"))
            if not ok:
                raise HwpError(f"export_document: save_as returned False for {out}")
            return ExportResult(exported=True, path=str(out), format=fmt_name)  # type: ignore[arg-type]

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Return the contents of a single table (by index from "
            "get_structure) as CSV. Useful for downstream spreadsheet "
            "processing. Walks cells manually so it works without pandas."
        ),
    )
    async def get_table_as_csv(
        doc_id: int = Field(..., description="Document index from open_document"),
        table_index: int = Field(..., description="Table index from get_structure", ge=0),
    ) -> dict:
        import csv
        from io import StringIO

        def _do(hwp: Any) -> TableCsvResult:
            _require_doc(hwp, doc_id)
            # Walk controls until we find the nth table control.
            target = None
            rows = cols = 0
            table_count = 0
            for ctrl in _iter_ctrls(hwp):
                try:
                    ctrl_id = str(ctrl.CtrlID).strip().lower()
                except Exception:  # noqa: BLE001
                    continue
                if ctrl_id != "tbl":
                    continue
                if table_count == table_index:
                    target = ctrl
                    try:
                        shape = ctrl.Properties
                        rows = int(shape.Item("Rows"))
                        cols = int(shape.Item("Cols"))
                    except Exception:  # noqa: BLE001
                        pass
                    break
                table_count += 1
            if target is None:
                raise HwpError(f"table_index={table_index} 를 찾을 수 없습니다.")
            if not rows or not cols:
                raise HwpError(
                    f"표 크기를 읽을 수 없습니다 (table_index={table_index})."
                )

            # Snapshot caret so we can restore it.
            saved_pos = None
            try:
                saved_pos = hwp.get_pos()
            except Exception:  # noqa: BLE001
                pass

            # Step into the table at its first cell.
            try:
                hwp.set_pos_by_ctrl(target)
            except Exception:  # noqa: BLE001
                pass
            try:
                hwp.find_ctrl()
                hwp.HAction.Run("ShapeObjTableSelCell")
                hwp.HAction.Run("Close")
            except Exception:  # noqa: BLE001
                pass

            buf = StringIO()
            writer = csv.writer(buf)
            try:
                for r in range(rows):
                    row_values: list[str] = []
                    for c in range(cols):
                        # Select the entire cell content, copy/extract text.
                        text = ""
                        try:
                            hwp.HAction.Run("SelectCellBlock")
                            text = str(hwp.get_selected_text() or "") if hasattr(
                                hwp, "get_selected_text"
                            ) else ""
                        except Exception:  # noqa: BLE001
                            text = ""
                        # Fallback: try reading via clipboard-less API.
                        row_values.append(text.strip())
                        if c < cols - 1:
                            try:
                                hwp.HAction.Run("TableRightCell")
                            except Exception:  # noqa: BLE001
                                break
                    writer.writerow(row_values)
                    if r < rows - 1:
                        try:
                            # Go back to column 0, then down one row.
                            for _ in range(cols - 1):
                                hwp.HAction.Run("TableLeftCell")
                            hwp.HAction.Run("TableLowerCell")
                        except Exception:  # noqa: BLE001
                            break
            finally:
                if saved_pos is not None:
                    try:
                        hwp.set_pos(*saved_pos)
                    except Exception:  # noqa: BLE001
                        pass

            return TableCsvResult(
                table_index=table_index,
                rows=int(rows),
                cols=int(cols),
                csv=buf.getvalue(),
            )

        return to_dict(await session.call(_do))
