"""Table helpers — cell address arithmetic and the ``fill_by_path`` DSL.

The DSL lets callers address a cell relative to a label in the document::

    "이름: > right"           # the cell immediately to the right of "이름:"
    "합계 > down > down"     # two cells below the "합계" label
    "Header > right > down"  # right one then down one

This keeps templates readable — no hard-coded (table_index, row, col) tuples
that break whenever a table gains a header row.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List

_DIRECTIONS = {"right", "left", "up", "down"}


@dataclass(frozen=True)
class TablePath:
    """A parsed ``fill_by_path`` expression."""

    label: str
    directions: tuple[str, ...]

    def apply(self, row: int, col: int) -> tuple[int, int]:
        """Return the (row, col) reached by applying each direction."""
        r, c = row, col
        for step in self.directions:
            if step == "right":
                c += 1
            elif step == "left":
                c -= 1
            elif step == "down":
                r += 1
            elif step == "up":
                r -= 1
        return r, c


def parse_path(expression: str) -> TablePath:
    """Parse ``"label > dir > dir"`` → :class:`TablePath`.

    Raises ``ValueError`` if the expression is malformed so the tool layer
    can surface a clean error instead of a silent mis-fill.
    """
    if not expression or not isinstance(expression, str):
        raise ValueError("path expression must be a non-empty string")

    parts: List[str] = [p.strip() for p in expression.split(">")]
    if not parts or not parts[0]:
        raise ValueError(f"path expression is missing a label: {expression!r}")

    label = parts[0]
    directions: list[str] = []
    for part in parts[1:]:
        if part.lower() not in _DIRECTIONS:
            raise ValueError(
                f"unknown direction {part!r} in path {expression!r}; "
                f"expected one of {sorted(_DIRECTIONS)}"
            )
        directions.append(part.lower())

    # A bare label with no directions makes no sense for fill_by_path —
    # the label itself is not a cell we want to overwrite. Require ≥1 step.
    if not directions:
        raise ValueError(
            f"path expression must include at least one direction: {expression!r}"
        )

    return TablePath(label=label, directions=tuple(directions))
