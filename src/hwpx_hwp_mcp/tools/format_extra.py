"""Extra formatting tools (category H).

Paragraph style, list style, multi-column layout, watermark, and document
properties.  These complement the character-level set_font tool in create.py.
"""

from __future__ import annotations

from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import AppliedResult, to_dict
from .session import _require_doc


_ALIGN_MAP = {
    "left": "Left",
    "center": "Center",
    "right": "Right",
    "justify": "Justify",
    "distribute": "Distribute",
    "distribute_space": "DistributeSpace",
}


def register(mcp: FastMCP) -> None:

    @mcp.tool(
        description=(
            "문단 모양 설정 / Set paragraph-level formatting on the current "
            "selection or the current paragraph. All parameters are optional; "
            "unset ones are left unchanged.\n\n"
            "- line_spacing: 줄 간격 percentage (100–500). 160 = 1.6×, 200 = 2×\n"
            "- space_before_pt / space_after_pt: 문단 위/아래 간격 (points)\n"
            "- indent_first_pt: 첫 줄 들여쓰기 (양수 = 들여쓰기, 음수 = 내어쓰기)\n"
            "- indent_left_pt / indent_right_pt: 왼쪽/오른쪽 여백 (points)\n"
            "- align: left | center | right | justify | distribute\n"
            "- page_break_before: 문단 앞 강제 쪽 나눔 여부\n"
            "- keep_lines: 문단 보호 (keep all lines on same page)\n"
            "- apply_to: 'selection' (default) | 'document' (entire document)"
        ),
    )
    async def set_paragraph_style(
        doc_id: int = Field(..., description="Document index from open_document"),
        line_spacing: Optional[int] = Field(
            None, description="Line spacing as percentage (100–500)", ge=100, le=500
        ),
        space_before_pt: Optional[float] = Field(
            None, description="Space before paragraph in points (0–841.8)", ge=0
        ),
        space_after_pt: Optional[float] = Field(
            None, description="Space after paragraph in points (0–841.8)", ge=0
        ),
        indent_first_pt: Optional[float] = Field(
            None,
            description=(
                "First-line indent in points. Positive = indent, "
                "negative = hanging indent."
            ),
        ),
        indent_left_pt: Optional[float] = Field(
            None, description="Left margin in points", ge=0
        ),
        indent_right_pt: Optional[float] = Field(
            None, description="Right margin in points", ge=0
        ),
        align: Optional[str] = Field(
            None,
            description="Paragraph alignment: left | center | right | justify | distribute",
        ),
        page_break_before: Optional[bool] = Field(
            None, description="Force a page break before this paragraph"
        ),
        keep_lines: Optional[bool] = Field(
            None, description="Prevent page breaks within the paragraph"
        ),
        apply_to: str = Field(
            "selection",
            description="'selection' = current paragraph(s); 'document' = entire document",
        ),
    ) -> dict:
        if align is not None and align.lower() not in _ALIGN_MAP:
            raise HwpError(
                f"align 값이 유효하지 않습니다: {align!r} "
                f"(허용: {sorted(_ALIGN_MAP)})"
            )
        if apply_to not in {"selection", "document"}:
            raise HwpError(
                f"apply_to 값이 유효하지 않습니다: {apply_to!r}"
            )

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            if apply_to == "document":
                hwp.SelectAll()

            kwargs: dict = {}
            if line_spacing is not None:
                kwargs["LineSpacing"] = int(line_spacing)
            if space_before_pt is not None:
                kwargs["PrevSpacing"] = float(space_before_pt)
            if space_after_pt is not None:
                kwargs["NextSpacing"] = float(space_after_pt)
            if indent_first_pt is not None:
                kwargs["Indentation"] = float(indent_first_pt)
            if indent_left_pt is not None:
                kwargs["LeftMargin"] = float(indent_left_pt)
            if indent_right_pt is not None:
                kwargs["RightMargin"] = float(indent_right_pt)
            if align is not None:
                kwargs["AlignType"] = _ALIGN_MAP[align.lower()]
            if page_break_before is not None:
                kwargs["PagebreakBefore"] = 1 if page_break_before else 0
            if keep_lines is not None:
                kwargs["KeepLinesTogether"] = 1 if keep_lines else 0

            if not kwargs:
                return AppliedResult(
                    applied=False, detail="no-op: 변경할 속성이 없습니다."
                )

            try:
                hwp.set_para(**kwargs)
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_para 실패: {exc}") from exc

            return AppliedResult(
                applied=True, detail=f"applied {sorted(kwargs)} to={apply_to}"
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "목록 스타일 적용 / Apply a bullet or numbered list style to the "
            "current paragraph(s). \n\n"
            "style_type options:\n"
            "- 'bullet' — •  bullet points\n"
            "- 'number' — 1. 2. 3. numbered list\n"
            "- 'none' — remove any existing list style\n\n"
            "For Korean-style outlines, use insert_paragraph with "
            "style='개요 1' / '개요 2' / '개요 3' instead."
        ),
    )
    async def set_list_style(
        doc_id: int = Field(..., description="Document index from open_document"),
        style_type: str = Field(
            "bullet",
            description="List style: bullet | number | none",
        ),
        apply_to: str = Field(
            "selection",
            description="'selection' = current paragraph(s); 'document' = entire document",
        ),
    ) -> dict:
        stype = style_type.strip().lower()
        action_map = {
            "bullet": "ParaBulletList",
            "number": "ParaNumList",
            "none": "ParaListOff",
        }
        if stype not in action_map:
            raise HwpError(
                f"style_type 이 유효하지 않습니다: {style_type!r} "
                f"(bullet / number / none)"
            )

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            if apply_to == "document":
                hwp.SelectAll()
            try:
                ok = bool(hwp.HAction.Run(action_map[stype]))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(
                    f"list style '{style_type}' 적용 실패: {exc}"
                ) from exc
            return AppliedResult(
                applied=ok, detail=f"list={style_type} apply_to={apply_to}"
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "다단 레이아웃 설정 / Set the number of text columns for the "
            "current section. columns=1 restores single-column layout. "
            "spacing_mm controls the gap between columns. "
            "Applies to the section containing the current caret position."
        ),
    )
    async def set_column_layout(
        doc_id: int = Field(..., description="Document index from open_document"),
        columns: int = Field(
            ..., description="Number of columns (1 = single column)", ge=1, le=10
        ),
        spacing_mm: float = Field(
            8.0, description="Gap between columns in millimeters", gt=0
        ),
        line_between: bool = Field(
            False, description="Draw a vertical line between columns"
        ),
        equal_width: bool = Field(
            True, description="Make all columns equal width"
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                # set_pagedef accepts MultiColCount and MultiColGap (mm→HWPUnit)
                page_args: dict = {
                    "MultiColCount": int(columns),
                    "MultiColGap": int(spacing_mm * 2835),
                }
                if not equal_width:
                    page_args["MultiColSameSize"] = 0
                ok = bool(hwp.set_pagedef(page_args))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"다단 설정 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=(
                    f"columns={columns} spacing={spacing_mm}mm "
                    f"line={line_between}"
                ),
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "워터마크 삽입 / Insert a text watermark on all pages of the "
            "document. The watermark is inserted as a background object. "
            "opacity is 0–100 (100 = fully opaque). angle is in degrees "
            "(315 = diagonal from bottom-left to top-right, default for '대외비')."
        ),
    )
    async def set_watermark(
        doc_id: int = Field(..., description="Document index from open_document"),
        text: str = Field(..., description="Watermark text, e.g. '대외비' or 'DRAFT'"),
        opacity: int = Field(
            30, description="Opacity 0–100 (30 = light watermark)", ge=0, le=100
        ),
        angle: int = Field(
            315, description="Rotation angle in degrees (315 = diagonal up-right)"
        ),
        font_size_pt: int = Field(
            60, description="Font size in points", ge=8, le=300
        ),
        color_hex: str = Field(
            "#C0C0C0", description="Text color as #RRGGBB hex string"
        ),
    ) -> dict:
        def _hex_to_int(h: str) -> int:
            v = h.lstrip("#").strip()
            r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
            return (b << 16) | (g << 8) | r  # HWP BGR packed int

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            color_int = _hex_to_int(color_hex)
            try:
                # Correct parameter set is HPrintWatermark
                pset = hwp.HParameterSet.HPrintWatermark
                hwp.HAction.GetDefault("PrintWatermark", pset.HSet)
                pset.string = str(text)                    # watermark text
                pset.WatermarkType = 1                     # 1 = text watermark
                pset.AlphaText = int(opacity * 255 // 100)  # 0–255
                pset.RotateAngle = int(angle)
                pset.FontSize = int(font_size_pt * 100)    # in 1/100 pt units
                pset.FontColor = color_int                 # BGR packed
                ok = bool(hwp.HAction.Execute("PrintWatermark", pset.HSet))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"워터마크 삽입 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"text={text!r} opacity={opacity} angle={angle}°",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "문서 속성 설정 / Set document metadata properties: title, author, "
            "subject, keywords, and description (comment). These are stored in "
            "the HWP file's document info and are visible in File→Properties. "
            "Pass only the fields you want to change; others are left as-is."
        ),
    )
    async def set_document_properties(
        doc_id: int = Field(..., description="Document index from open_document"),
        title: Optional[str] = Field(None, description="Document title"),
        author: Optional[str] = Field(None, description="Author name"),
        subject: Optional[str] = Field(None, description="Subject / topic"),
        keywords: Optional[str] = Field(
            None, description="Keywords (comma-separated)"
        ),
        description: Optional[str] = Field(
            None, description="Description / comment"
        ),
    ) -> dict:
        if all(
            v is None for v in [title, author, subject, keywords, description]
        ):
            raise HwpError("설정할 속성이 하나도 지정되지 않았습니다.")

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            applied: list[str] = []
            try:
                # Document metadata is exposed via IXHwpDocument.XHwpSummaryInfo
                doc = hwp.XHwpDocuments.Item(doc_id)
                summary = doc.XHwpSummaryInfo
                if title is not None:
                    summary.Title = str(title)
                    applied.append("title")
                if author is not None:
                    summary.Author = str(author)
                    applied.append("author")
                if subject is not None:
                    summary.Subject = str(subject)
                    applied.append("subject")
                if keywords is not None:
                    summary.Keywords = str(keywords)
                    applied.append("keywords")
                if description is not None:
                    summary.Comments = str(description)
                    applied.append("description")
                ok = bool(applied)
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"문서 속성 설정 실패: {exc}") from exc
            return AppliedResult(
                applied=ok, detail=f"updated: {applied}"
            )

        return to_dict(await session.call(_do))
