"""Regression smoke test for ``insert_table(data=...)``.

Reproduces the user-reported bug where a 5x10 table created via
``insert_table`` with explicit data left the saved file in a state that
한/글 reported as "파일이 손상되었습니다" on reopen. The root cause was
calling ``HAction.Run("CloseEx")`` after the table inserts — that action
closes the active *document*, not the table edit mode.

This script exercises the full path: create empty doc → insert table with
data → save_as → close → reopen → verify each cell value is present in
the extracted text.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / "table_with_data.hwpx"

    rows = 5
    cols = 10
    data = [[f"R{r}C{c}" for c in range(cols)] for r in range(rows)]
    expected_cells = {cell for row in data for cell in row}

    def _write(hwp):
        hwp.add_doc()
        # Mirror what tools/create.py:insert_table does when data is given.
        ok = hwp.create_table(
            rows=rows, cols=cols, treat_as_char=False, header=True
        )
        assert ok, "create_table returned False"
        for r_idx, row in enumerate(data):
            for c_idx in range(cols):
                value = row[c_idx]
                if value:
                    hwp.insert_text(value)
                if not (r_idx == rows - 1 and c_idx == cols - 1):
                    hwp.HAction.Run("TableRightCell")
        saved = hwp.save_as(str(target), format="HWPX")
        assert saved, "save_as returned False"
        return str(target)

    wrote = session.call_sync(_write)
    print(f"[1/2] wrote: {wrote}")

    def _read(hwp):
        ok = hwp.open(str(target), format="", arg="lock:true")
        assert ok, "reopen returned False"
        return hwp.get_text_file("TEXT", "") or ""

    text = session.call_sync(_read)
    print(f"[2/2] read back {len(text)} chars")

    missing = sorted(c for c in expected_cells if c not in text)
    if missing:
        raise SystemExit(
            f"FAIL: {len(missing)}/{len(expected_cells)} cells missing. "
            f"first few: {missing[:10]}"
        )
    print(f"OK - all {len(expected_cells)} cell values present after reopen")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
