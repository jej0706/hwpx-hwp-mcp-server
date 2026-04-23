"""Editing control tools (category F).

Low-level caret / selection / undo-redo tools that give the LLM fine-grained
control over the editing state of the active document.
"""

from __future__ import annotations

from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import AppliedResult, to_dict
from .session import _require_doc


def register(mcp: FastMCP) -> None:

    @mcp.tool(
        description=(
            "실행 취소 / Undo the last N editing actions in the document. "
            "Equivalent to pressing Ctrl+Z. Use count > 1 to undo multiple "
            "steps in a single call."
        ),
    )
    async def undo(
        doc_id: int = Field(..., description="Document index from open_document"),
        count: int = Field(1, description="Number of undo steps (1–50)", ge=1, le=50),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            for _ in range(count):
                hwp.HAction.Run("Undo")
            return AppliedResult(applied=True, detail=f"undo ×{count}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "다시 실행 / Redo the last N undone actions. "
            "Equivalent to pressing Ctrl+Y."
        ),
    )
    async def redo(
        doc_id: int = Field(..., description="Document index from open_document"),
        count: int = Field(1, description="Number of redo steps (1–50)", ge=1, le=50),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            for _ in range(count):
                hwp.HAction.Run("Redo")
            return AppliedResult(applied=True, detail=f"redo ×{count}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "한/글 HAction 직접 실행 (탈출구) / Execute any HWP action by its "
            "string ID. This is the universal escape hatch for any HWP feature "
            "not covered by other tools.\n\n"
            "Commonly useful action IDs:\n"
            "- 'SelectAll' — select entire document\n"
            "- 'MoveDocBegin' / 'MoveDocEnd' — jump to start / end\n"
            "- 'MoveNextParaBegin' / 'MovePrevParaBegin' — paragraph hop\n"
            "- 'MoveParaBegin' / 'MoveParaEnd' — line start / end\n"
            "- 'Cancel' — cancel selection / exit sub-area (header, table, etc.)\n"
            "- 'CloseEx' — close sub-area and return to body\n"
            "- 'BreakPara' — insert paragraph break\n"
            "- 'Delete' — delete selected content\n"
            "- 'Copy' / 'Cut' / 'Paste' — clipboard ops\n"
            "⚠️ Some actions open modal dialogs that block the server thread."
        ),
    )
    async def run_action(
        doc_id: int = Field(..., description="Document index from open_document"),
        action_id: str = Field(
            ..., description="HWP action ID string, e.g. 'SelectAll'"
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.HAction.Run(str(action_id)))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(
                    f"HAction.Run('{action_id}') 실패: {exc}"
                ) from exc
            return AppliedResult(applied=ok, detail=f"action={action_id!r}")

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "현재 선택된 텍스트 반환 / Return the text currently selected in "
            "the document. If nothing is selected and the caret is inside a "
            "table cell, returns the cell's text. If the caret is in body "
            "text with no selection, returns the current word. "
            "Use select_text first to set a deliberate selection."
        ),
    )
    async def get_selection_text(
        doc_id: int = Field(..., description="Document index from open_document"),
        keep_selection: bool = Field(
            True,
            description="Keep the selection active after reading (default True)",
        ),
    ) -> dict:
        def _do(hwp: Any) -> dict:
            _require_doc(hwp, doc_id)
            try:
                text = hwp.get_selected_text(
                    as_="str", keep_select=keep_selection
                )
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"get_selected_text 실패: {exc}") from exc
            return {"text": text or "", "length": len(text or "")}

        return await session.call(_do)

    @mcp.tool(
        description=(
            "텍스트 범위 선택 / Select a range of text by paragraph (0-based) "
            "and character offset (0-based). The epos character is NOT "
            "included in the selection (half-open interval). Use "
            "get_document_text to discover paragraph structure. "
            "After selecting, call get_selection_text to confirm, or "
            "run_action('Copy') to copy the selection to the clipboard.\n\n"
            "Example: select the whole first paragraph → "
            "start_para=0, start_pos=0, end_para=0, end_pos=-1"
        ),
    )
    async def select_text(
        doc_id: int = Field(..., description="Document index from open_document"),
        start_para: int = Field(
            ..., description="Start paragraph index (0-based)", ge=0
        ),
        start_pos: int = Field(
            0, description="Start char offset within the paragraph (0-based)", ge=0
        ),
        end_para: int = Field(
            ..., description="End paragraph index (0-based)", ge=0
        ),
        end_pos: int = Field(
            -1,
            description=(
                "End char offset (-1 means end of the paragraph). "
                "The character AT this offset is NOT included."
            ),
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            # pyhwpx select_text: epos=-1 not directly supported; use large
            # number to simulate "end of paragraph".
            actual_epos = end_pos if end_pos >= 0 else 99_999
            try:
                ok = bool(
                    hwp.select_text(
                        spara=start_para,
                        spos=start_pos,
                        epara=end_para,
                        epos=actual_epos,
                        slist=0,
                    )
                )
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"select_text 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=(
                    f"para {start_para}:{start_pos} → "
                    f"{end_para}:{end_pos}"
                ),
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "현재 캐럿 위치 반환 / Get the current caret position as "
            "(list_id, para, pos). list_id=0 means body text; other values "
            "indicate sub-lists (table cells, header/footer, footnotes, etc.). "
            "para and pos are both 0-based."
        ),
    )
    async def get_caret_pos(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> dict:
            _require_doc(hwp, doc_id)
            try:
                list_id, para, pos = hwp.get_pos()
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"get_pos 실패: {exc}") from exc
            return {"list_id": int(list_id), "para": int(para), "pos": int(pos)}

        return await session.call(_do)

    @mcp.tool(
        description=(
            "캐럿 위치 이동 / Move the caret to a specific paragraph and "
            "character offset in the document body (list_id=0). "
            "para and pos are 0-based; pos=-1 moves to end of the paragraph."
        ),
    )
    async def set_caret_pos(
        doc_id: int = Field(..., description="Document index from open_document"),
        para: int = Field(..., description="Paragraph index (0-based)", ge=0),
        pos: int = Field(
            0,
            description="Character offset within the paragraph (0-based; -1 = end)",
        ),
        list_id: int = Field(
            0, description="List ID (0 = body text; use get_caret_pos to discover others)"
        ),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                ok = bool(hwp.set_pos(List=list_id, para=para, pos=pos))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"set_pos 실패: {exc}") from exc
            return AppliedResult(
                applied=ok, detail=f"list={list_id} para={para} pos={pos}"
            )

        return to_dict(await session.call(_do))
