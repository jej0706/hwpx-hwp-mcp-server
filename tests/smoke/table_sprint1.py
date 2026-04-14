"""End-to-end smoke test for Sprint 1 Section A table tools.

Exercises: merge_cells, split_cell, set_column_width, set_row_height,
set_cell_alignment, insert_table_row, delete_table_row,
insert_table_column, delete_table_column, set_cell_border.

The test drives the tools in the same sequence Claude would when composing
a stylized table, then saves and reopens the file to make sure it remains
valid HWPX.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / "table_sprint1.hwpx"
    if target.exists():
        target.unlink()

    def setup(hwp):
        hwp.add_doc()
        hwp.create_table(rows=4, cols=5, treat_as_char=False, header=True)
        for r in range(4):
            for c in range(5):
                hwp.insert_text(f"{r + 1},{c + 1}")
                if not (r == 3 and c == 4):
                    hwp.HAction.Run("TableRightCell")

    session.call_sync(setup)
    print("[setup] created 4x5 table with labels")

    # --- merge_cells ---
    def op_merge(hwp):
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A1")
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        # extend the block two cells to the right so A1:C1 is selected
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("TableRightCell")
        return bool(hwp.TableMergeCell())

    assert session.call_sync(op_merge), "merge failed"
    print("[1/10] merge_cells A1:C1 ok")

    # --- split_cell ---
    def op_split(hwp):
        hwp.get_into_nth_table(0)
        # After the merge, the first row has 3 columns (A/B/C). Split B1
        # into a 2x2 grid.
        hwp.goto_addr("B1")
        return bool(hwp.TableSplitCell(Rows=2, Cols=2, DistributeHeight=1, Merge=0))

    assert session.call_sync(op_split), "split failed"
    print("[2/10] split_cell B1 -> 2x2 ok")

    # --- set_column_width ---
    def op_colwidth(hwp):
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A1")
        return bool(hwp.set_col_width(30.0, as_="mm"))

    assert session.call_sync(op_colwidth), "set_col_width failed"
    print("[3/10] set_column_width A=30mm ok")

    # --- set_row_height ---
    def op_rowheight(hwp):
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A2")
        return bool(hwp.set_row_height(12.0, as_="mm"))

    assert session.call_sync(op_rowheight), "set_row_height failed"
    print("[4/10] set_row_height row2=12mm ok")

    # --- set_cell_alignment (header row center/center) ---
    def op_align(hwp):
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A1")
        hwp.TableCellBlock()
        hwp.TableCellBlockRow()
        return bool(hwp.HAction.Run("TableCellAlignCenterCenter"))

    assert session.call_sync(op_align), "set_cell_alignment failed"
    print("[5/10] set_cell_alignment row1=center/center ok")

    # --- set_cell_border (outside frame on all cells) ---
    def op_border(hwp):
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A1")
        hwp.TableCellBlock()
        hwp.TableCellBlockCol()
        hwp.TableCellBlockRow()
        return bool(hwp.HAction.Run("TableCellBorderOutside"))

    assert session.call_sync(op_border), "set_cell_border failed"
    print("[6/10] set_cell_border outside ok")

    # --- insert_table_row (append at end) ---
    def op_inserted_row(hwp):
        hwp.get_into_nth_table(0)
        return bool(hwp.TableAppendRow())

    assert session.call_sync(op_inserted_row), "insert_table_row failed"
    print("[7/10] insert_table_row append ok")

    # --- insert_table_column (append at right) ---
    def op_inserted_col(hwp):
        hwp.get_into_nth_table(0)
        try:
            n = int(hwp.get_col_num())
        except Exception:
            n = 0
        if n:
            letter = chr(ord("A") + n - 1)
            hwp.goto_addr(f"{letter}1")
        return bool(hwp.TableRightCellAppend())

    assert session.call_sync(op_inserted_col), "insert_table_column failed"
    print("[8/10] insert_table_column append ok")

    # --- delete_table_row (remove the last row we just added) ---
    def op_delete_row(hwp):
        hwp.get_into_nth_table(0)
        try:
            n = int(hwp.get_row_num())
        except Exception:
            n = 5
        hwp.goto_addr(f"A{n}")
        hwp.TableCellBlock()
        hwp.TableCellBlockRow()
        return bool(hwp.TableDeleteCell(remain_cell=False))

    assert session.call_sync(op_delete_row), "delete_table_row failed"
    print("[9/10] delete_table_row last ok")

    # --- delete_table_column (remove the last column we just added) ---
    def op_delete_col(hwp):
        hwp.get_into_nth_table(0)
        try:
            n = int(hwp.get_col_num())
        except Exception:
            n = 6
        letter = chr(ord("A") + n - 1)
        hwp.goto_addr(f"{letter}1")
        hwp.TableCellBlock()
        hwp.TableCellBlockCol()
        return bool(hwp.TableDeleteCell(remain_cell=False))

    assert session.call_sync(op_delete_col), "delete_table_column failed"
    print("[10/10] delete_table_column last ok")

    # --- save and reopen ---
    def op_save(hwp):
        return bool(hwp.save_as(str(target), format="HWPX"))

    assert session.call_sync(op_save), "save_as returned False"
    assert target.exists(), f"file not written: {target}"
    print(f"[save] {target} ({target.stat().st_size} bytes)")

    def op_reopen(hwp):
        ok = hwp.open(str(target), format="", arg="lock:true")
        assert ok, "reopen returned False"
        return hwp.get_text_file("TEXT", "") or ""

    text = session.call_sync(op_reopen)
    print(f"[reopen] text length {len(text)}")
    # After merge/split/insert/delete the exact label set is unpredictable,
    # but the file must reopen and retain SOME of the cell labels.
    surviving = sum(1 for r in range(1, 5) for c in range(1, 6) if f"{r},{c}" in text)
    assert surviving >= 10, (
        f"too few cell labels survived: {surviving}/20 (file likely corrupt)"
    )
    print(f"OK - Sprint 1 Section A smoke test passed ({surviving}/20 labels kept)")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
