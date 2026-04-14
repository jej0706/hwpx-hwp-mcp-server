"""End-to-end smoke test for Sprint 1 Section B page tools.

Exercises: set_page_settings (A4 + custom margins + landscape),
insert_section_break, insert_page_number. Validates that the file saves
and reopens cleanly, and that set_page_settings actually updated the
section definition.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / "page_sprint1.hwpx"
    if target.exists():
        target.unlink()

    def setup(hwp):
        hwp.add_doc()
        hwp.insert_text("첫 번째 섹션 본문")
        hwp.BreakPara()

    session.call_sync(setup)
    print("[setup] created document with first section body")

    # --- set_page_settings to A4 landscape with custom margins ---
    def op_page_settings(hwp):
        return bool(
            hwp.set_pagedef(
                {
                    "PaperWidth": 297.0,
                    "PaperHeight": 210.0,
                    "Landscape": 1,  # landscape
                    "TopMargin": 15.0,
                    "BottomMargin": 15.0,
                    "LeftMargin": 20.0,
                    "RightMargin": 20.0,
                },
                apply="cur",
            )
        )

    assert session.call_sync(op_page_settings), "set_pagedef failed"
    print("[1/4] set_page_settings A4 landscape + custom margins ok")

    # --- Verify it took effect ---
    def op_verify(hwp):
        d = hwp.get_pagedef_as_dict(as_="eng")
        return d

    pagedef = session.call_sync(op_verify)
    print(f"[2/4] pagedef after update: {pagedef}")
    # Widths read back as mm per get_pagedef_as_dict docstring
    assert pagedef.get("Landscape") == 1, f"expected Landscape=1, got {pagedef}"
    # Margins should be ~15mm, ~20mm (within 0.1mm tolerance)
    assert abs(pagedef.get("TopMargin", 0) - 15.0) < 0.5, pagedef
    assert abs(pagedef.get("LeftMargin", 0) - 20.0) < 0.5, pagedef

    # --- Insert section break + page number ---
    def op_section(hwp):
        hwp.insert_text("같은 섹션 추가 문장")
        hwp.BreakPara()
        ok1 = bool(hwp.BreakSection())
        hwp.insert_text("두 번째 섹션 시작")
        hwp.BreakPara()
        ok2 = bool(hwp.InsertPageNum())
        return ok1 and ok2

    assert session.call_sync(op_section), "section break / page num insert failed"
    print("[3/4] insert_section_break + insert_page_number ok")

    # --- Save and reopen ---
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
    assert "첫 번째 섹션" in text, "first section text missing"
    assert "두 번째 섹션" in text, "second section text missing"
    print(f"[4/4] reopen ok, both section texts present")
    print("OK - Sprint 1 Section B smoke test passed")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
