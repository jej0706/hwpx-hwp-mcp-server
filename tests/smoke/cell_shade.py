"""Regression smoke test for ``set_cell_shade``.

Reproduces the user-reported scenario: create a 5x10 table with header row
shaded gray, save as HWPX, reopen, verify it still loads. The original bug
was that Claude bypassed the MCP entirely (because there was no shading
tool) and edited HWPX XML directly, producing a corrupt file. Adding the
``set_cell_shade`` tool removes the incentive to bypass.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / "shaded_header_5x10.hwpx"
    if target.exists():
        target.unlink()

    rows, cols = 5, 10

    def _create(hwp):
        hwp.add_doc()
        return hwp.XHwpDocuments.Count - 1

    doc_id = session.call_sync(_create)

    def _fill_table(hwp):
        hwp.switch_to(doc_id)
        hwp.create_table(rows=rows, cols=cols, treat_as_char=False, header=True)
        for r in range(rows):
            for c in range(cols):
                hwp.insert_text(f"{r + 1},{c + 1}")
                if not (r == rows - 1 and c == cols - 1):
                    hwp.HAction.Run("TableRightCell")

    session.call_sync(_fill_table)

    def _shade_header(hwp):
        hwp.switch_to(doc_id)
        # Mirror what set_cell_shade(cells="row:1") does internally.
        hwp.get_into_nth_table(0)
        hwp.goto_addr("A1")
        hwp.TableCellBlock()
        hwp.TableCellBlockRow()
        return bool(hwp.cell_fill((217, 217, 217)))

    shaded = session.call_sync(_shade_header)
    print(f"[1/3] header row shaded: {shaded}")

    def _save(hwp):
        hwp.switch_to(doc_id)
        return bool(hwp.save_as(str(target), format="HWPX"))

    saved = session.call_sync(_save)
    assert saved, "save_as returned False"
    assert target.exists(), f"file not written: {target}"
    print(f"[2/3] saved: {target} ({target.stat().st_size} bytes)")

    def _reopen(hwp):
        ok = hwp.open(str(target), format="", arg="lock:true")
        assert ok, "reopen returned False"
        return hwp.get_text_file("TEXT", "") or ""

    text = session.call_sync(_reopen)
    print(f"[3/3] reopened, text length: {len(text)}")

    # Spot-check a few cell values to confirm content was preserved
    for label in ("1,1", "1,10", "5,1", "5,10"):
        assert label in text, f"missing cell {label} in reopened doc"

    print("OK - cell shade smoke test passed (file opens cleanly + content intact)")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
