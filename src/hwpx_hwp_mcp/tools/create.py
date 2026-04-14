"""Creation tools (category D).

These tools write *new* content into a document. They are intentionally
high-level - callers describe "insert a paragraph aligned centre" instead
of running raw HAction commands - so Claude can compose documents from
scratch with a small number of tool calls.
"""

from __future__ import annotations

from typing import Any, List, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import AppliedResult, InsertResult, InsertTableResult, to_dict
from ..utils.paths import ensure_existing_file
from .session import _require_doc


_ALIGN_MAP = {
    "left": "Left",
    "center": "Center",
    "right": "Right",
    "justify": "Justify",
    "distribute": "Distribute",
}


def _hex_to_rgb_tuple(color_hex: str) -> tuple[int, int, int]:
    """Parse ``"#RRGGBB"`` (or ``"RRGGBB"``) into an ``(r, g, b)`` tuple."""
    v = color_hex.lstrip("#").strip()
    if len(v) != 6:
        raise HwpError(
            f"color_hex 형식이 잘못되었습니다: {color_hex!r} (예: '#D9D9D9')"
        )
    try:
        return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    except ValueError as exc:
        raise HwpError(f"color_hex 파싱 실패: {color_hex!r}") from exc


def _col_number_to_letter(n: int) -> str:
    """1-based column number to Excel letter. 1→A, 26→Z, 27→AA."""
    if n < 1:
        raise HwpError(f"열 번호는 1 이상이어야 합니다: {n}")
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def _letter_to_col_number(letter: str) -> int:
    """Excel letter to 1-based column number. A→1, Z→26, AA→27."""
    letter = letter.strip().upper()
    if not letter or not letter.isalpha():
        raise HwpError(f"알파벳 열 지정이 잘못되었습니다: {letter!r}")
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _enter_table(hwp: Any, table_index: int) -> None:
    """Move the caret into the Nth table (0-based) or raise."""
    try:
        ok = hwp.get_into_nth_table(table_index)
    except Exception as exc:  # noqa: BLE001
        raise HwpError(
            f"table_index={table_index} 로 이동할 수 없습니다: {exc}"
        ) from exc
    if ok is False:
        raise HwpError(f"table_index={table_index} 를 찾을 수 없습니다.")


def _parse_addr(addr: str) -> tuple[int, int]:
    """Parse an Excel-style address (``'B3'``) into (row, col) 1-based."""
    s = addr.strip().upper()
    if not s:
        raise HwpError("셀 주소가 비어있습니다.")
    letters = ""
    digits = ""
    for ch in s:
        if ch.isalpha():
            if digits:
                raise HwpError(f"셀 주소 형식 오류: {addr!r}")
            letters += ch
        elif ch.isdigit():
            digits += ch
        else:
            raise HwpError(f"셀 주소 형식 오류: {addr!r}")
    if not letters or not digits:
        raise HwpError(f"셀 주소에 열 또는 행이 누락: {addr!r}")
    return int(digits), _letter_to_col_number(letters)


def _extend_block_to(hwp: Any, right: int, down: int) -> None:
    """Extend the current cell block by ``right`` columns and ``down`` rows.

    Negative values move left / up instead. Caller must have already entered
    cell-block mode via :func:`_select_cells` or directly.
    """
    action_right = "TableRightCell" if right >= 0 else "TableLeftCell"
    action_down = "TableLowerCell" if down >= 0 else "TableUpperCell"
    for _ in range(abs(int(right))):
        hwp.HAction.Run(action_right)
    for _ in range(abs(int(down))):
        hwp.HAction.Run(action_down)


def _select_cells(hwp: Any, cells: str) -> None:
    """Apply a cells selector to the current table.

    Supports ``"all"``, ``"row:N"``, ``"col:N"`` / ``"col:L"``, Excel-style
    single addresses ``"A1"``, and rectangular ranges ``"A1:C3"``. Caller
    must have already moved the caret into the target table via
    :func:`_enter_table`.

    The selection uses the "cell block + extend" pattern. ``TableCellBlock``
    alone only picks the current cell — you must call ``TableCellBlockExtend``
    before directional moves (``TableRightCell``/``TableLowerCell``) actually
    extend the selection instead of moving the caret.
    """
    spec = (cells or "all").strip()
    spec_lower = spec.lower()

    if spec_lower == "all":
        try:
            n_cols = int(hwp.get_col_num())
            n_rows = int(hwp.get_row_num())
        except Exception as exc:  # noqa: BLE001
            raise HwpError(
                f"표 크기를 조회할 수 없어 'all' 선택이 실패했습니다: {exc}"
            ) from exc
        hwp.goto_addr("A1")
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        _extend_block_to(hwp, right=max(0, n_cols - 1), down=max(0, n_rows - 1))
        return

    if spec_lower.startswith("row:"):
        try:
            n = int(spec.split(":", 1)[1])
        except ValueError as exc:
            raise HwpError(f"row:N 의 N 은 정수여야 합니다: {cells!r}") from exc
        if n < 1:
            raise HwpError(f"row 번호는 1 이상이어야 합니다: {n}")
        hwp.goto_addr(f"A{n}")
        hwp.TableCellBlock()
        hwp.TableCellBlockRow()
        return

    if spec_lower.startswith("col:"):
        tail = spec.split(":", 1)[1].strip()
        if tail.isalpha():
            addr = f"{tail.upper()}1"
        else:
            try:
                n = int(tail)
            except ValueError as exc:
                raise HwpError(
                    f"col:N 의 N 은 정수 또는 알파벳이어야 합니다: {cells!r}"
                ) from exc
            addr = f"{_col_number_to_letter(n)}1"
        hwp.goto_addr(addr)
        hwp.TableCellBlock()
        hwp.TableCellBlockCol()
        return

    if ":" in spec:
        start_addr, end_addr = (s.strip() for s in spec.split(":", 1))
        if not start_addr or not end_addr:
            raise HwpError(f"범위 형식이 잘못되었습니다: {cells!r}")
        r1, c1 = _parse_addr(start_addr)
        r2, c2 = _parse_addr(end_addr)
        hwp.goto_addr(start_addr)
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        _extend_block_to(hwp, right=c2 - c1, down=r2 - r1)
        return

    # Single cell, Excel notation
    hwp.goto_addr(spec)
    hwp.TableCellBlock()


_HORIZONTAL_ALIGN = {"left": "Left", "center": "Center", "right": "Right"}
_VERTICAL_ALIGN = {"top": "Top", "center": "Center", "bottom": "Bottom"}


# Map a user-friendly ``sides`` value to the pyhwpx TableCellBorder* action.
_BORDER_SIDES = {
    "all": "TableCellBorderAll",
    "outside": "TableCellBorderOutside",
    "inside": "TableCellBorderInside",
    "inside_horz": "TableCellBorderInsideHorz",
    "inside_vert": "TableCellBorderInsideVert",
    "top": "TableCellBorderTop",
    "bottom": "TableCellBorderBottom",
    "left": "TableCellBorderLeft",
    "right": "TableCellBorderRight",
    "diagonal_down": "TableCellBorderDiagonalDown",
    "diagonal_up": "TableCellBorderDiagonalUp",
    "none": "TableCellBorderNo",
}


def register(mcp: FastMCP) -> None:
    @mcp.tool(
        description=(
            "Insert a paragraph of text at the current caret position. "
            "Always appends a paragraph break after the text, so consecutive "
            "calls produce consecutive paragraphs. Optionally applies a named "
            "built-in style (e.g. '제목 1') and an alignment."
        ),
    )
    async def insert_paragraph(
        doc_id: int = Field(..., description="Document index from open_document"),
        text: str = Field(..., description="Paragraph text (may be empty for a blank line)"),
        style: Optional[str] = Field(
            None, description="Optional named style to apply to the paragraph"
        ),
        align: Optional[str] = Field(
            None,
            description="Optional alignment: left | center | right | justify | distribute",
        ),
    ) -> dict:
        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            if align is not None and align not in _ALIGN_MAP:
                raise HwpError(f"align 값이 유효하지 않습니다: {align!r}")
            if style:
                try:
                    hwp.set_style(style)
                except Exception as exc:  # noqa: BLE001
                    raise HwpError(f"스타일 '{style}' 적용 실패: {exc}") from exc
            if align:
                try:
                    hwp.set_para(AlignType=_ALIGN_MAP[align])
                except Exception as exc:  # noqa: BLE001
                    raise HwpError(f"정렬 '{align}' 적용 실패: {exc}") from exc
            if text:
                hwp.insert_text(text)
            hwp.BreakPara()
            return InsertResult(inserted=True, detail=None)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Insert a table at the current caret position. If `data` is "
            "provided it is written row-by-row; otherwise an empty table "
            "is created. `header=True` makes the first row a header."
        ),
    )
    async def insert_table(
        doc_id: int = Field(..., description="Document index from open_document"),
        rows: int = Field(..., description="Number of rows (ignored when data is given)", ge=1),
        cols: int = Field(..., description="Number of columns (ignored when data is given)", ge=1),
        data: Optional[List[List[str]]] = Field(
            None,
            description=(
                "Optional 2D array of string cell values. If given, rows/cols "
                "are inferred from its shape."
            ),
        ),
        header: bool = Field(True, description="Treat the first row as a header"),
    ) -> dict:
        def _do(hwp: Any) -> InsertTableResult:
            _require_doc(hwp, doc_id)
            if data is not None and data:
                actual_rows = len(data)
                actual_cols = max(len(r) for r in data)
                # Do NOT use pyhwpx.table_from_data - it requires real pandas.
                # Create the shell, then walk the cells and insert_text into
                # each one. This is a little slower but keeps us pandas-free.
                ok = hwp.create_table(
                    rows=actual_rows,
                    cols=actual_cols,
                    treat_as_char=False,
                    header=header,
                )
                if not ok:
                    raise HwpError("create_table 가 실패했습니다.")
                # After create_table the caret lands in the first cell.
                # IMPORTANT: do NOT call HAction.Run("CloseEx") afterwards -
                # CloseEx closes the active *document*, not the table edit
                # mode, and corrupts the file we are about to save. Leave the
                # caret wherever it ends up; callers can navigate explicitly
                # if they need to insert content after the table.
                for r_idx, row in enumerate(data):
                    for c_idx in range(actual_cols):
                        value = str(row[c_idx]) if c_idx < len(row) else ""
                        if value:
                            hwp.insert_text(value)
                        # Advance to the next cell unless we are on the last one.
                        if not (r_idx == actual_rows - 1 and c_idx == actual_cols - 1):
                            try:
                                hwp.HAction.Run("TableRightCell")
                            except Exception:  # noqa: BLE001
                                break
            else:
                ok = hwp.create_table(
                    rows=rows, cols=cols, treat_as_char=False, header=header
                )
                if not ok:
                    raise HwpError("create_table 가 실패했습니다.")
                actual_rows, actual_cols = rows, cols
            return InsertTableResult(
                inserted=True, rows=int(actual_rows), cols=int(actual_cols)
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Insert an image at the current caret position. image_path must "
            "be an absolute Windows path. width/height are in HWPUNIT (pass 0 "
            "to keep the image's native size). as_char=True anchors the "
            "image like a character; False anchors it to the paragraph."
        ),
    )
    async def insert_image(
        doc_id: int = Field(..., description="Document index from open_document"),
        image_path: str = Field(..., description="Absolute path to the image file"),
        width: int = Field(0, description="Width in HWPUNIT; 0 = native", ge=0),
        height: int = Field(0, description="Height in HWPUNIT; 0 = native", ge=0),
        as_char: bool = Field(True, description="Anchor as a character (vs. paragraph)"),
        embedded: bool = Field(True, description="Embed the image file into the document"),
    ) -> dict:
        resolved = ensure_existing_file(image_path)

        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            ctrl = hwp.insert_picture(
                str(resolved),
                treat_as_char=as_char,
                embedded=embedded,
                width=width,
                height=height,
            )
            return InsertResult(
                inserted=ctrl is not None,
                detail=str(resolved),
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description="Insert a hard page break at the current caret position.",
    )
    async def insert_page_break(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            ok = bool(hwp.BreakPage())
            return InsertResult(inserted=ok)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Change character formatting on the current selection (or the "
            "whole document when apply_to='document'). Any parameter left as "
            "None is unchanged. Colors are hex strings like '#FF0000'."
        ),
    )
    async def set_font(
        doc_id: int = Field(..., description="Document index from open_document"),
        family: Optional[str] = Field(None, description="Font family / face name"),
        size_pt: Optional[float] = Field(None, description="Font size in points", gt=0),
        bold: Optional[bool] = Field(None, description="Bold on/off"),
        italic: Optional[bool] = Field(None, description="Italic on/off"),
        underline: Optional[bool] = Field(None, description="Underline on/off"),
        color_hex: Optional[str] = Field(
            None, description="Text color as #RRGGBB hex string"
        ),
        apply_to: str = Field(
            "selection",
            description="selection | document - the range to apply the change to",
        ),
    ) -> dict:
        if apply_to not in {"selection", "document"}:
            raise HwpError(f"apply_to 값이 유효하지 않습니다: {apply_to!r}")

        def _hex_to_int(h: str) -> int:
            r, g, b = _hex_to_rgb_tuple(h)
            # HWP stores color as BGR packed int.
            return (b << 16) | (g << 8) | r

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            if apply_to == "document":
                hwp.SelectAll()

            kwargs: dict = {}
            if family is not None:
                kwargs["FaceName"] = family
            if size_pt is not None:
                # pyhwpx set_font uses Height in points per its signature.
                kwargs["Height"] = size_pt
            if bold is not None:
                kwargs["Bold"] = bool(bold)
            if italic is not None:
                kwargs["Italic"] = bool(italic)
            if underline is not None:
                kwargs["UnderlineType"] = 1 if underline else 0
            if color_hex is not None:
                kwargs["TextColor"] = _hex_to_int(color_hex)

            if not kwargs:
                return AppliedResult(applied=False, detail="no-op: 변경할 속성이 없습니다.")

            try:
                hwp.set_font(**kwargs)
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_font 실패: {exc}") from exc
            return AppliedResult(applied=True, detail=", ".join(sorted(kwargs)))

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표 셀 배경색(음영) 지정 / Set table cell background color. "
            "셀 색칠, 배경색 채우기, 헤더 행 음영, cell shade / fill / "
            "background color - 이 도구를 사용하세요. "
            "⚠️ **절대 HWPX XML 을 직접 편집하지 마세요** - XML 을 직접 "
            "생성하면 OWPML 스키마 불일치로 파일이 '손상됨' 상태가 됩니다. "
            "이 도구가 한/글 COM 의 CellFill 액션을 호출해서 안전하게 "
            "처리합니다.\n\n"
            "cells 셀렉터:\n"
            "- 'all': 표 전체\n"
            "- 'row:N' (1-based): N 번째 행 전체 (예: 'row:1' = 헤더 행)\n"
            "- 'col:N' 또는 'col:L' (1-based 숫자 또는 알파벳): 열 전체\n"
            "- 'A1' (Excel 표기): 단일 셀\n"
            "- 'A1:C3' (Excel 표기): 직사각형 범위\n\n"
            "color_hex 기본값은 #D9D9D9 (연회색, 헤더용). "
            "예: 하늘색 = #87CEEB, 빨강 = #FF0000, 연노랑 = #FFFFCC."
        ),
    )
    async def set_cell_shade(
        doc_id: int = Field(..., description="Document index from open_document"),
        cells: str = Field(
            "all",
            description=(
                "Cell selector. One of: 'all' | 'row:N' | 'col:N' | 'A1' | "
                "'A1:C3'. Row/col numbers are 1-based; addresses use Excel "
                "notation."
            ),
        ),
        color_hex: str = Field(
            "#D9D9D9",
            description="Fill color as #RRGGBB hex string (default light gray)",
        ),
        table_index: int = Field(
            0,
            description="Which table to target (0-based, in document order)",
            ge=0,
        ),
    ) -> dict:
        rgb = _hex_to_rgb_tuple(color_hex)

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            _select_cells(hwp, cells)
            try:
                ok = bool(hwp.cell_fill(rgb))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"cell_fill 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} cells={cells} color={color_hex}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표 셀 병합 / Merge selected cells in a table. "
            "cells 에는 반드시 범위(예: 'A1:C1', 'B2:B4')를 주세요. "
            "단일 셀('A1')이나 'all' 도 받지만 의미 있는 병합이 되려면 "
            "2개 이상의 셀을 포함해야 합니다."
        ),
    )
    async def merge_cells(
        doc_id: int = Field(..., description="Document index from open_document"),
        cells: str = Field(
            ...,
            description=(
                "Cell range to merge. Use Excel range notation like 'A1:C1' "
                "(merge first three cells of row 1) or 'B2:B4' (merge three "
                "rows of column B). Also accepts 'row:N' or 'col:N'."
            ),
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            _select_cells(hwp, cells)
            try:
                ok = bool(hwp.TableMergeCell())
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableMergeCell 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} merged={cells}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "셀 분할 / Split a single cell into a grid. Positions the caret "
            "at the given single-cell address (e.g. 'B2'), then splits that "
            "cell into rows x cols sub-cells. Cannot split a range."
        ),
    )
    async def split_cell(
        doc_id: int = Field(..., description="Document index from open_document"),
        cell: str = Field(..., description="Single cell address in Excel notation (e.g. 'B2')"),
        rows: int = Field(..., description="Number of rows after split", ge=1),
        cols: int = Field(..., description="Number of columns after split", ge=1),
        distribute_height: bool = Field(
            False, description="Distribute heights evenly after split"
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        if rows == 1 and cols == 1:
            raise HwpError("rows 와 cols 둘 다 1 이면 분할 의미가 없습니다.")

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            hwp.goto_addr(cell.strip())
            try:
                ok = bool(
                    hwp.TableSplitCell(
                        Rows=int(rows),
                        Cols=int(cols),
                        DistributeHeight=1 if distribute_height else 0,
                        Merge=0,
                    )
                )
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableSplitCell 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} cell={cell} -> {rows}x{cols}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "열 너비 지정 / Set the width of a single table column in "
            "millimeters. `col` can be an Excel letter ('A', 'B', ...) or "
            "a 1-based integer. Width is in mm."
        ),
    )
    async def set_column_width(
        doc_id: int = Field(..., description="Document index from open_document"),
        col: str = Field(
            ..., description="Column letter ('A') or 1-based number ('1')"
        ),
        width_mm: float = Field(..., description="Column width in millimeters", gt=0),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        col_str = str(col).strip()
        if col_str.isdigit():
            col_letter = _col_number_to_letter(int(col_str))
        else:
            col_letter = col_str.upper()
            _letter_to_col_number(col_letter)

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            hwp.goto_addr(f"{col_letter}1")
            try:
                ok = bool(hwp.set_col_width(float(width_mm), as_="mm"))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_col_width 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} col={col_letter} width={width_mm}mm",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "행 높이 지정 / Set the height of a single table row in "
            "millimeters. `row` is a 1-based integer."
        ),
    )
    async def set_row_height(
        doc_id: int = Field(..., description="Document index from open_document"),
        row: int = Field(..., description="1-based row number", ge=1),
        height_mm: float = Field(..., description="Row height in millimeters", gt=0),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            hwp.goto_addr(f"A{row}")
            try:
                ok = bool(hwp.set_row_height(float(height_mm), as_="mm"))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_row_height 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} row={row} height={height_mm}mm",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "셀 정렬 / Set text alignment inside table cells. `horizontal` "
            "is one of 'left' | 'center' | 'right'; `vertical` is one of "
            "'top' | 'center' | 'bottom'. The `cells` selector accepts the "
            "same values as set_cell_shade."
        ),
    )
    async def set_cell_alignment(
        doc_id: int = Field(..., description="Document index from open_document"),
        cells: str = Field(
            "all",
            description="Cell selector: 'all' | 'row:N' | 'col:N' | 'A1' | 'A1:C3'",
        ),
        horizontal: str = Field(
            "center", description="Horizontal: left | center | right"
        ),
        vertical: str = Field(
            "center", description="Vertical: top | center | bottom"
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        h = horizontal.strip().lower()
        v = vertical.strip().lower()
        if h not in _HORIZONTAL_ALIGN:
            raise HwpError(
                f"horizontal 값이 유효하지 않습니다: {horizontal!r} (left/center/right)"
            )
        if v not in _VERTICAL_ALIGN:
            raise HwpError(
                f"vertical 값이 유효하지 않습니다: {vertical!r} (top/center/bottom)"
            )
        action_name = f"TableCellAlign{_HORIZONTAL_ALIGN[h]}{_VERTICAL_ALIGN[v]}"

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            _select_cells(hwp, cells)
            try:
                ok = bool(hwp.HAction.Run(action_name))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"{action_name} 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} cells={cells} align={h}/{v}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표에 행 추가 / Append a new row to the table. If `at_row` is "
            "given (1-based), the new row is inserted BELOW that row. "
            "Otherwise the row is appended at the end of the table."
        ),
    )
    async def insert_table_row(
        doc_id: int = Field(..., description="Document index from open_document"),
        at_row: int = Field(
            0,
            description=(
                "1-based row number; new row is inserted BELOW this row. "
                "Use 0 to append at the end of the table."
            ),
            ge=0,
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            if at_row > 0:
                hwp.goto_addr(f"A{at_row}")
            try:
                ok = bool(hwp.TableAppendRow())
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableAppendRow 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} inserted below row {at_row or 'last'}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표 행 삭제 / Delete a row from a table. Positions the caret "
            "at the given 1-based row number, selects it, then removes it."
        ),
    )
    async def delete_table_row(
        doc_id: int = Field(..., description="Document index from open_document"),
        row: int = Field(..., description="1-based row number to delete", ge=1),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            hwp.goto_addr(f"A{row}")
            hwp.TableCellBlock()
            hwp.TableCellBlockRow()
            try:
                ok = bool(hwp.TableDeleteCell(remain_cell=False))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableDeleteCell 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} deleted row {row}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표에 열 추가 / Append a new column to the right of the table. "
            "Uses TableRightCellAppend which adds one column at the right edge."
        ),
    )
    async def insert_table_column(
        doc_id: int = Field(..., description="Document index from open_document"),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            try:
                n_cols = int(hwp.get_col_num())
            except Exception:  # noqa: BLE001
                n_cols = 0
            if n_cols:
                hwp.goto_addr(f"{_col_number_to_letter(n_cols)}1")
            try:
                ok = bool(hwp.TableRightCellAppend())
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableRightCellAppend 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} column appended",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "용지 크기/방향/여백 설정 / Configure page size, orientation, "
            "and margins for the current section. All margin values are in "
            "millimeters. Missing fields are left unchanged.\n\n"
            "- paper: A3 | A4 | A5 | B4 | B5 | Letter | Legal (or 'custom' "
            "with paper_width_mm/paper_height_mm)\n"
            "- orientation: 'portrait' | 'landscape'\n"
            "- apply_to: 'cur' (current section only) | 'all' | 'new'"
        ),
    )
    async def set_page_settings(
        doc_id: int = Field(..., description="Document index from open_document"),
        paper: Optional[str] = Field(
            None,
            description="Paper preset: A3 | A4 | A5 | B4 | B5 | Letter | Legal | custom",
        ),
        orientation: Optional[str] = Field(
            None, description="portrait | landscape"
        ),
        top_mm: Optional[float] = Field(None, description="Top margin (mm)", ge=0),
        bottom_mm: Optional[float] = Field(None, description="Bottom margin (mm)", ge=0),
        left_mm: Optional[float] = Field(None, description="Left margin (mm)", ge=0),
        right_mm: Optional[float] = Field(None, description="Right margin (mm)", ge=0),
        header_mm: Optional[float] = Field(None, description="Header height (mm)", ge=0),
        footer_mm: Optional[float] = Field(None, description="Footer height (mm)", ge=0),
        paper_width_mm: Optional[float] = Field(
            None, description="Custom paper width in mm (used when paper='custom')", gt=0
        ),
        paper_height_mm: Optional[float] = Field(
            None, description="Custom paper height in mm (used when paper='custom')", gt=0
        ),
        apply_to: str = Field(
            "cur", description="Scope: 'cur' | 'all' | 'new'"
        ),
    ) -> dict:
        # Paper presets in mm. pyhwpx set_pagedef converts mm to HWPUnit.
        presets = {
            "a3": (297, 420),
            "a4": (210, 297),
            "a5": (148, 210),
            "b4": (257, 364),
            "b5": (182, 257),
            "letter": (215.9, 279.4),
            "legal": (215.9, 355.6),
        }
        if apply_to not in {"cur", "all", "new"}:
            raise HwpError(f"apply_to 값이 유효하지 않습니다: {apply_to!r}")

        dict_args: dict[str, float] = {}

        if paper is not None:
            key = paper.strip().lower()
            if key == "custom":
                if not (paper_width_mm and paper_height_mm):
                    raise HwpError(
                        "paper='custom' 일 때는 paper_width_mm 와 paper_height_mm 가 필요합니다."
                    )
                dict_args["PaperWidth"] = float(paper_width_mm)
                dict_args["PaperHeight"] = float(paper_height_mm)
            elif key in presets:
                w, h = presets[key]
                dict_args["PaperWidth"] = w
                dict_args["PaperHeight"] = h
            else:
                raise HwpError(
                    f"paper 값이 유효하지 않습니다: {paper!r} (허용: {sorted(presets)} 또는 'custom')"
                )
        elif paper_width_mm and paper_height_mm:
            dict_args["PaperWidth"] = float(paper_width_mm)
            dict_args["PaperHeight"] = float(paper_height_mm)

        if orientation is not None:
            ori = orientation.strip().lower()
            if ori == "portrait":
                dict_args["Landscape"] = 0
            elif ori == "landscape":
                dict_args["Landscape"] = 1
            else:
                raise HwpError(
                    f"orientation 값이 유효하지 않습니다: {orientation!r}"
                )

        if top_mm is not None:
            dict_args["TopMargin"] = float(top_mm)
        if bottom_mm is not None:
            dict_args["BottomMargin"] = float(bottom_mm)
        if left_mm is not None:
            dict_args["LeftMargin"] = float(left_mm)
        if right_mm is not None:
            dict_args["RightMargin"] = float(right_mm)
        if header_mm is not None:
            dict_args["HeaderLen"] = float(header_mm)
        if footer_mm is not None:
            dict_args["FooterLen"] = float(footer_mm)

        if not dict_args:
            raise HwpError("설정할 항목이 없습니다. 최소 하나의 파라미터를 지정해주세요.")

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.set_pagedef(dict_args, apply=apply_to))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_pagedef 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"{sorted(dict_args)} apply_to={apply_to}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "현재 위치에 쪽 번호 필드 삽입 / Insert a page-number field at "
            "the current caret position. Uses Hancom's default page-number "
            "format and position. Call this inside a header/footer area for "
            "headers/footers (requires a header/footer to already exist in "
            "the document)."
        ),
    )
    async def insert_page_number(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.InsertPageNum())
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"InsertPageNum 실패: {exc}") from exc
            return AppliedResult(applied=ok, detail="page number inserted at caret")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "구역 나누기 삽입 / Insert a section break at the current caret "
            "position. A new section inherits most settings from the "
            "previous one but can have different page orientation, "
            "header/footer, column layout, etc. Use this before calling "
            "set_page_settings(apply_to='cur') to change layout for a "
            "portion of the document."
        ),
    )
    async def insert_section_break(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.BreakSection())
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"BreakSection 실패: {exc}") from exc
            return AppliedResult(applied=ok, detail="section break inserted")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표 셀 테두리 / Apply cell borders to selected cells using the "
            "current default line style. `sides` picks which edges to "
            "toggle: 'all' (every edge), 'outside' (outer frame only), "
            "'inside' (inner lines only), 'top'/'bottom'/'left'/'right' for "
            "a single edge, 'inside_horz'/'inside_vert' for inner horizontal "
            "or vertical only, 'diagonal_down'/'diagonal_up' for diagonals, "
            "or 'none' to remove borders from the selection.\n\n"
            "⚠️ Border thickness and color are NOT customizable through this "
            "tool in v0.1 — Hancom applies whatever the current default line "
            "style is. For thick or colored borders, set the border via HWP "
            "GUI once and then re-use through this tool, or wait for a "
            "future update."
        ),
    )
    async def set_cell_border(
        doc_id: int = Field(..., description="Document index from open_document"),
        cells: str = Field(
            "all",
            description="Cell selector: 'all' | 'row:N' | 'col:N' | 'A1' | 'A1:C3'",
        ),
        sides: str = Field(
            "all",
            description=(
                "Which edges to toggle: all | outside | inside | top | bottom | "
                "left | right | inside_horz | inside_vert | diagonal_down | "
                "diagonal_up | none"
            ),
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        sides_key = (sides or "all").strip().lower()
        if sides_key not in _BORDER_SIDES:
            raise HwpError(
                f"sides 값이 유효하지 않습니다: {sides!r} "
                f"(허용: {sorted(_BORDER_SIDES)})"
            )
        action_name = _BORDER_SIDES[sides_key]

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            _select_cells(hwp, cells)
            try:
                ok = bool(hwp.HAction.Run(action_name))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"{action_name} 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} cells={cells} sides={sides_key}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "표 열 삭제 / Delete a column from a table. `col` can be an "
            "Excel letter or 1-based integer."
        ),
    )
    async def delete_table_column(
        doc_id: int = Field(..., description="Document index from open_document"),
        col: str = Field(
            ..., description="Column letter ('A') or 1-based number ('1')"
        ),
        table_index: int = Field(
            0, description="Which table to target (0-based)", ge=0
        ),
    ) -> dict:
        col_str = str(col).strip()
        if col_str.isdigit():
            col_letter = _col_number_to_letter(int(col_str))
        else:
            col_letter = col_str.upper()
            _letter_to_col_number(col_letter)

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            _enter_table(hwp, table_index)
            hwp.goto_addr(f"{col_letter}1")
            hwp.TableCellBlock()
            hwp.TableCellBlockCol()
            try:
                ok = bool(hwp.TableDeleteCell(remain_cell=False))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"TableDeleteCell 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"table={table_index} deleted col {col_letter}",
            )

        return to_dict(await session.call(_do))
