"""Template-filling tools (category C).

These are the highest-value tools for the user's main use case: taking a
pre-made 한/글 양식 and filling in field values without disturbing its
layout or fonts.

Two filling strategies are exposed:

- ``fill_fields`` — writes into 누름틀(field) by name via
  ``put_field_text``. Preserves all surrounding formatting. Use when the
  template already has named fields.
- ``replace_text`` — scans the whole document for a literal or regex and
  substitutes it. Use when the template has `{{name}}`-style placeholders
  instead of real fields.

``fill_table_by_path`` supports a lightweight DSL so templates that label
cells (e.g. "이름:") can be filled without hardcoding table indices.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import (
    CreateFieldResult,
    FillFieldsResult,
    FillTablePathResult,
    ReplaceTextResult,
    to_dict,
)
from ..utils.tables import parse_path
from .read import _split_field_list
from .session import _require_doc


def register(mcp: FastMCP) -> None:
    @mcp.tool(
        description=(
            "List every 누름틀(field) defined in the document, with the "
            "current text inside each one. Always call this before "
            "fill_fields so you know which names exist."
        ),
    )
    async def list_fields(
        doc_id: int = Field(..., description="Document index from open_document"),
    ) -> dict:
        def _do(hwp: Any) -> dict:
            _require_doc(hwp, doc_id)
            raw = ""
            try:
                raw = hwp.get_field_list(1, 0)
            except Exception:  # noqa: BLE001
                pass
            names = _split_field_list(raw)
            out: List[dict] = []
            for i, name in enumerate(names):
                current: Optional[str] = None
                try:
                    current = str(hwp.get_field_text(name))
                except Exception:  # noqa: BLE001
                    pass
                out.append({"name": name, "index": i, "current_text": current})
            return {"fields": out, "count": len(out)}

        return await session.call(_do)

    @mcp.tool(
        description=(
            "Fill one or more 누름틀 fields in the active document. Preserves "
            "the original font and paragraph styling — prefer this over "
            "replace_text when a template has real fields. To target a "
            "specific instance of a repeated field use name {{0}} / {{1}} "
            "(e.g. '고객명{{1}}')."
        ),
    )
    async def fill_fields(
        doc_id: int = Field(..., description="Document index from open_document"),
        values: Dict[str, str] = Field(
            ...,
            description="Mapping of field name → text value. Keys may include {{N}} suffix.",
        ),
    ) -> dict:
        def _do(hwp: Any) -> FillFieldsResult:
            _require_doc(hwp, doc_id)
            existing = set(_split_field_list(hwp.get_field_list(1, 0)))
            unknown: List[str] = []
            # Strip the {{N}} suffix when validating against existing names.
            for raw_name in values.keys():
                base = raw_name.split("{{", 1)[0]
                if base not in existing:
                    unknown.append(raw_name)

            # pyhwpx's put_field_text accepts dict directly.
            hwp.put_field_text(values, "")
            return FillFieldsResult(filled=len(values), unknown_fields=unknown)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Create a new 누름틀 field at the current caret position (or at a "
            "specified list/para/pos). Useful when adding placeholders to a "
            "document that does not yet have them."
        ),
    )
    async def create_field(
        doc_id: int = Field(..., description="Document index from open_document"),
        name: str = Field(..., description="Name for the new field"),
        list_idx: Optional[int] = Field(
            None,
            description="Optional list index to target before inserting (part of get_pos tuple)",
        ),
        para: Optional[int] = Field(None, description="Optional paragraph index"),
        pos: Optional[int] = Field(None, description="Optional character position within the paragraph"),
    ) -> dict:
        def _do(hwp: Any) -> CreateFieldResult:
            _require_doc(hwp, doc_id)
            if list_idx is not None and para is not None and pos is not None:
                hwp.set_pos(list_idx, para, pos)
            try:
                hwp.create_field(name, direction="", memo="")
            except TypeError:
                # Older pyhwpx variants accept (name,) only.
                hwp.create_field(name)
            return CreateFieldResult(created=True, name=name)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Find and replace text across the whole document. Set regex=True "
            "to interpret `old` as a regular expression. Use this when the "
            "template uses literal placeholders like '{{name}}' instead of "
            "real 누름틀 fields."
        ),
    )
    async def replace_text(
        doc_id: int = Field(..., description="Document index from open_document"),
        old: str = Field(..., description="Text to search for"),
        new: str = Field(..., description="Replacement text"),
        regex: bool = Field(False, description="Interpret `old` as regex"),
        all: bool = Field(True, description="Replace every occurrence (default) or just the next one"),
    ) -> dict:
        def _do(hwp: Any) -> ReplaceTextResult:
            _require_doc(hwp, doc_id)
            if all:
                count = hwp.find_replace_all(old, new, regex=regex)
            else:
                count = 1 if hwp.find_replace(old, new, regex=regex) else 0
            try:
                count = int(count)
            except Exception:  # noqa: BLE001
                count = int(bool(count))
            return ReplaceTextResult(replaced=count)

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Fill table cells using a label-relative DSL. Each key is a path "
            "like '이름: > right' or '합계 > down > down'. The label is found "
            "in the document, the direction steps move the caret, and the "
            "value is written into the resulting cell. Great for filling "
            "forms that pair labels and blank cells."
        ),
    )
    async def fill_table_by_path(
        doc_id: int = Field(..., description="Document index from open_document"),
        mappings: Dict[str, str] = Field(
            ..., description="path expression → value (see tool description)"
        ),
    ) -> dict:
        def _do(hwp: Any) -> FillTablePathResult:
            _require_doc(hwp, doc_id)
            filled = 0
            misses: List[str] = []

            for expression, value in mappings.items():
                try:
                    path = parse_path(expression)
                except ValueError as exc:
                    misses.append(f"{expression}: {exc}")
                    continue

                # Snapshot cursor so a miss does not leave it dangling.
                saved = None
                try:
                    saved = hwp.get_pos()
                except Exception:  # noqa: BLE001
                    pass
                # Jump to document start before searching so the label is
                # found deterministically.
                try:
                    hwp.MovePos(2)
                except Exception:  # noqa: BLE001
                    pass

                if not hwp.find_forward(path.label, regex=False):
                    misses.append(f"{expression}: label not found")
                    if saved is not None:
                        try:
                            hwp.set_pos(*saved)
                        except Exception:  # noqa: BLE001
                            pass
                    continue

                moved_ok = True
                for step in path.directions:
                    action = {
                        "right": "TableRightCell",
                        "left": "TableLeftCell",
                        "down": "TableLowerCell",
                        "up": "TableUpperCell",
                    }[step]
                    try:
                        if not hwp.HAction.Run(action):
                            moved_ok = False
                            break
                    except Exception:  # noqa: BLE001
                        moved_ok = False
                        break
                if not moved_ok:
                    misses.append(f"{expression}: direction navigation failed")
                    if saved is not None:
                        try:
                            hwp.set_pos(*saved)
                        except Exception:  # noqa: BLE001
                            pass
                    continue

                try:
                    # Clear the target cell and type the new value.
                    hwp.HAction.Run("SelectCellBlock")
                    hwp.HAction.Run("Delete")
                    hwp.insert_text(value)
                    filled += 1
                except Exception as exc:  # noqa: BLE001
                    misses.append(f"{expression}: write failed ({exc})")

                if saved is not None:
                    try:
                        hwp.set_pos(*saved)
                    except Exception:  # noqa: BLE001
                        pass

            return FillTablePathResult(filled=filled, misses=misses)

        return to_dict(await session.call(_do))
