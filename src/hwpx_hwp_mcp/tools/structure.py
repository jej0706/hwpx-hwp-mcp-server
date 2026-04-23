"""Document structure tools (category G).

Tools that insert structural elements: headers/footers, footnotes, bookmarks,
hyperlinks, table of contents, text boxes, and shapes.
"""

from __future__ import annotations

from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import AppliedResult, InsertResult, to_dict
from .session import _require_doc


def register(mcp: FastMCP) -> None:

    @mcp.tool(
        description=(
            "머리말/꼬리말 삽입·수정 / Insert or replace the header and/or "
            "footer text for the current section. Pass None to skip either "
            "part. If a header/footer already exists in the section, its "
            "content is replaced; otherwise a new one is created.\n\n"
            "Tip: Call save_document after setting header/footer to persist "
            "the changes before inspecting."
        ),
    )
    async def insert_header_footer(
        doc_id: int = Field(..., description="Document index from open_document"),
        header_text: Optional[str] = Field(
            None, description="Text for the page header (None = leave unchanged)"
        ),
        footer_text: Optional[str] = Field(
            None, description="Text for the page footer (None = leave unchanged)"
        ),
        align: str = Field(
            "center",
            description="Text alignment inside the header/footer: left | center | right",
        ),
    ) -> dict:
        if header_text is None and footer_text is None:
            raise HwpError(
                "header_text 와 footer_text 중 최소 하나는 지정해야 합니다."
            )

        align_map = {"left": "Left", "center": "Center", "right": "Right"}
        if align.lower() not in align_map:
            raise HwpError(
                f"align 값이 유효하지 않습니다: {align!r} (left/center/right)"
            )
        align_type = align_map[align.lower()]

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            applied: list[str] = []

            for ctrl_id, text in (("head", header_text), ("foot", footer_text)):
                if text is None:
                    continue
                try:
                    existing = hwp.get_ctrl_by_ctrl_id(ctrl_id)
                    if existing:
                        # Navigate to anchor, then enter the sub-list
                        ctrl = existing[0]
                        pos = hwp.get_ctrl_pos(ctrl)
                        hwp.set_pos(*pos)
                        # Move into the header/footer sub-list via F6 equivalent
                        hwp.HAction.Run("MoveNextPosEx")
                        hwp.SelectAll()
                        # Set alignment before inserting text
                        try:
                            hwp.set_para(AlignType=align_type)
                        except Exception:  # noqa: BLE001
                            pass
                        if text:
                            hwp.insert_text(text)
                        hwp.HAction.Run("CloseEx")
                    else:
                        # Create new header/footer via parameterized action
                        pset = hwp.HParameterSet.HHeaderFooter
                        hwp.HAction.GetDefault("HeaderFooter", pset.HSet)
                        pset.HSet.SetItem(
                            "Type", 0 if ctrl_id == "head" else 1
                        )
                        hwp.HAction.Execute("HeaderFooter", pset.HSet)
                        # Cursor should now be inside the new header/footer
                        hwp.SelectAll()
                        try:
                            hwp.set_para(AlignType=align_type)
                        except Exception:  # noqa: BLE001
                            pass
                        if text:
                            hwp.insert_text(text)
                        hwp.HAction.Run("CloseEx")
                    applied.append(ctrl_id)
                except Exception as exc:  # noqa: BLE001
                    raise HwpError(
                        f"{'머리말' if ctrl_id == 'head' else '꼬리말'} 처리 실패: {exc}"
                    ) from exc

            return AppliedResult(
                applied=bool(applied),
                detail=f"updated: {applied} align={align_type}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "각주/미주 삽입 / Insert a footnote (or endnote) at the current "
            "caret position. The caret returns to the body text after the "
            "footnote is inserted. Text is placed inside the footnote area."
        ),
    )
    async def insert_footnote(
        doc_id: int = Field(..., description="Document index from open_document"),
        text: str = Field(..., description="Footnote content text"),
        is_endnote: bool = Field(
            False, description="True = endnote (미주) instead of footnote (각주)"
        ),
    ) -> dict:
        action = "InsertEndnote" if is_endnote else "InsertFootnote"
        label = "미주" if is_endnote else "각주"

        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.HAction.Run(action))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"{label} 생성 실패: {exc}") from exc
            if not ok:
                raise HwpError(f"HAction.Run('{action}') returned False")
            # Cursor is now inside the footnote area — insert the text
            try:
                hwp.insert_text(text)
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"{label} 텍스트 삽입 실패: {exc}") from exc
            # Return to main body
            hwp.HAction.Run("CloseEx")
            return InsertResult(inserted=True, detail=f"{label}: {text[:40]!r}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "책갈피 삽입 / Insert a named bookmark at the current caret "
            "position. Bookmarks can be jumped to via hyperlinks or used as "
            "anchor targets for cross-references. Bookmark names must be "
            "unique within the document."
        ),
    )
    async def insert_bookmark(
        doc_id: int = Field(..., description="Document index from open_document"),
        name: str = Field(
            ...,
            description=(
                "Bookmark name (alphanumeric + underscore recommended, "
                "unique within the document)"
            ),
        ),
    ) -> dict:
        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            try:
                pset = hwp.HParameterSet.HFieldCtrl
                hwp.HAction.GetDefault("InsertFieldCtrl", pset.HSet)
                pset.HSet.SetItem("CtrlID", "%bmk")
                pset.HSet.SetItem("Name", str(name))
                ok = bool(hwp.HAction.Execute("InsertFieldCtrl", pset.HSet))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"책갈피 삽입 실패: {exc}") from exc
            return InsertResult(inserted=ok, detail=f"bookmark={name!r}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "하이퍼링크 삽입 / Insert a hyperlink on the currently selected "
            "text. If nothing is selected, the URL is inserted as link text. "
            "Two link types:\n"
            "- URL hyperlink: set url='https://...', leave bookmark_name empty\n"
            "- Internal bookmark link: set bookmark_name='my_bookmark', "
            "leave url empty\n"
            "Select the text you want to hyperlink first, then call this tool."
        ),
    )
    async def insert_hyperlink(
        doc_id: int = Field(..., description="Document index from open_document"),
        url: Optional[str] = Field(
            None, description="Destination URL (for external links)"
        ),
        bookmark_name: Optional[str] = Field(
            None,
            description="Bookmark name inside the document (for internal links)",
        ),
        display_text: Optional[str] = Field(
            None,
            description=(
                "Text to display for the link. If None, the currently "
                "selected text is used (or the URL/bookmark name)."
            ),
        ),
        tooltip: Optional[str] = Field(
            None, description="Tooltip text shown on hover"
        ),
    ) -> dict:
        if not url and not bookmark_name:
            raise HwpError("url 또는 bookmark_name 중 하나는 지정해야 합니다.")

        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)

            # If display_text is specified, insert it and select it first.
            if display_text:
                hwp.insert_text(display_text)
                # Select the just-inserted text
                try:
                    hwp.select_text(
                        spara=hwp.get_pos()[1],
                        spos=max(0, hwp.get_pos()[2] - len(display_text)),
                        epara=hwp.get_pos()[1],
                        epos=hwp.get_pos()[2],
                    )
                except Exception:  # noqa: BLE001
                    pass

            try:
                if url:
                    # URL hyperlink via HHyperLink parameter set
                    pset = hwp.HParameterSet.HHyperLink
                    hwp.HAction.GetDefault("InsertHyperlink", pset.HSet)
                    desc = tooltip or ""
                    # Command format for URL: "u<URL>|<desc>;..."
                    pset.Command = f"u{url}|{desc};0;0;0;"
                    ok = bool(hwp.HAction.Execute("InsertHyperlink", pset.HSet))
                else:
                    # Bookmark hyperlink
                    ok = bool(
                        hwp.insert_hyperlink(
                            hypertext=str(bookmark_name),
                            description=tooltip or "",
                        )
                    )
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"하이퍼링크 삽입 실패: {exc}") from exc

            target = url or bookmark_name
            return InsertResult(inserted=ok, detail=f"link → {target!r}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "목차 삽입 / Insert a Table of Contents at the current caret "
            "position. The TOC is built from headings (제목 1/2/3 styles) "
            "already in the document. If no headings exist yet, insert them "
            "first using insert_paragraph with style='제목 1'."
        ),
    )
    async def insert_toc(
        doc_id: int = Field(..., description="Document index from open_document"),
        levels: int = Field(
            3, description="Number of heading levels to include (1–4)", ge=1, le=4
        ),
    ) -> dict:
        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            try:
                pset = hwp.HParameterSet.HMakeToc
                hwp.HAction.GetDefault("MakeToc", pset.HSet)
                # Set the depth of heading levels to include
                try:
                    pset.HSet.SetItem("Levels", int(levels))
                except Exception:  # noqa: BLE001
                    pass
                ok = bool(hwp.HAction.Execute("MakeToc", pset.HSet))
            except Exception:  # noqa: BLE001
                # Fallback: try the simple run (may open dialog)
                try:
                    ok = bool(hwp.HAction.Run("MakeToc"))
                except Exception as exc2:  # noqa: BLE001
                    raise HwpError(f"목차 삽입 실패: {exc2}") from exc2
            return InsertResult(inserted=ok, detail=f"TOC levels={levels}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "글상자 삽입 / Insert a floating text box at the current caret "
            "position. The box is created with the given dimensions; optional "
            "text is placed inside it. Width and height are in millimeters."
        ),
    )
    async def insert_text_box(
        doc_id: int = Field(..., description="Document index from open_document"),
        text: Optional[str] = Field(
            None, description="Initial text content inside the text box"
        ),
        width_mm: float = Field(
            80.0, description="Box width in millimeters", gt=0
        ),
        height_mm: float = Field(
            40.0, description="Box height in millimeters", gt=0
        ),
        as_char: bool = Field(
            False,
            description=(
                "True = anchor as inline character; "
                "False = free-floating anchor (default)"
            ),
        ),
    ) -> dict:
        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            try:
                # Convert mm to HWPUnit (1 mm ≈ 2835 HWPUnit)
                w_hu = int(width_mm * 2835)
                h_hu = int(height_mm * 2835)

                # Use insert_ctrl("gso") which reliably inserts shape objects
                gso_pset = hwp.create_set("ShapeObject")
                try:
                    gso_pset.SetItem("ShapeType", 3)       # 3 = TextBox
                    gso_pset.SetItem("TreatAsChar", 1 if as_char else 0)
                    gso_pset.SetItem("Width", w_hu)
                    gso_pset.SetItem("Height", h_hu)
                except Exception:  # noqa: BLE001
                    pass
                ctrl = hwp.insert_ctrl("gso", gso_pset)
                ok = ctrl is not None
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"글상자 삽입 실패: {exc}") from exc

            # Enter the text box and insert text
            if ok and text:
                try:
                    hwp.ShapeObjTextBoxEdit()
                    hwp.insert_text(text)
                    hwp.HAction.Run("CloseEx")
                except Exception as exc:  # noqa: BLE001
                    raise HwpError(
                        f"글상자 텍스트 삽입 실패: {exc}"
                    ) from exc

            return InsertResult(
                inserted=ok,
                detail=f"{width_mm}×{height_mm}mm as_char={as_char}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "도형 삽입 / Insert a drawing shape at the current caret position. "
            "shape_type selects the shape:\n"
            "- 'rectangle' (사각형)\n"
            "- 'ellipse' (타원)\n"
            "- 'line' (선)\n"
            "- 'rounded_rectangle' (둥근 사각형)\n"
            "Width and height are in millimeters. "
            "fill_color_hex sets the fill (e.g. '#87CEEB'); None = no fill."
        ),
    )
    async def insert_shape(
        doc_id: int = Field(..., description="Document index from open_document"),
        shape_type: str = Field(
            "rectangle",
            description="Shape type: rectangle | ellipse | line | rounded_rectangle",
        ),
        width_mm: float = Field(40.0, description="Shape width in mm", gt=0),
        height_mm: float = Field(20.0, description="Shape height in mm", gt=0),
        fill_color_hex: Optional[str] = Field(
            None, description="Fill color as #RRGGBB hex (None = no fill)"
        ),
    ) -> dict:
        shape_map = {
            "rectangle": "DrawingObjectLine",       # placeholder action prefix
            "ellipse": "DrawingObjectEllipse",
            "line": "DrawingObjectLine",
            "rounded_rectangle": "DrawingObjectRoundRect",
        }
        # Map friendly names to pyhwpx / HWP InsertDrawingObject type values
        # Type: 0=Line, 1=Rectangle, 2=Ellipse, 3=TextBox, 4=RoundRect
        type_map = {
            "line": 0,
            "rectangle": 1,
            "ellipse": 2,
            "rounded_rectangle": 4,
        }
        stype = shape_type.lower().strip()
        if stype not in type_map:
            raise HwpError(
                f"shape_type 이 유효하지 않습니다: {shape_type!r} "
                f"(허용: rectangle / ellipse / line / rounded_rectangle)"
            )

        def _hex_to_int(h: str) -> int:
            v = h.lstrip("#").strip()
            r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
            return (b << 16) | (g << 8) | r  # BGR packed

        def _do(hwp: Any) -> InsertResult:
            _require_doc(hwp, doc_id)
            w_hu = int(width_mm * 2835)
            h_hu = int(height_mm * 2835)
            try:
                # Use insert_ctrl("gso") with ShapeType set
                gso_pset = hwp.create_set("ShapeObject")
                try:
                    gso_pset.SetItem("ShapeType", type_map[stype])
                    gso_pset.SetItem("Width", w_hu)
                    gso_pset.SetItem("Height", h_hu)
                    gso_pset.SetItem("TreatAsChar", 0)
                    if fill_color_hex:
                        gso_pset.SetItem("FillColor", _hex_to_int(fill_color_hex))
                except Exception:  # noqa: BLE001
                    pass
                ctrl = hwp.insert_ctrl("gso", gso_pset)
                ok = ctrl is not None
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"도형 삽입 실패: {exc}") from exc
            # Exit shape editing if caret entered shape sub-list
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:  # noqa: BLE001
                pass
            return InsertResult(
                inserted=ok,
                detail=f"shape={stype} {width_mm}×{height_mm}mm",
            )

        return to_dict(await session.call(_do))
