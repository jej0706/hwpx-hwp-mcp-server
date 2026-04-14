"""Manual smoke test: create → insert content → save → reopen → verify.

Run with Hancom HWP installed::

    python tests/smoke/cycle.py

This exercises the end-to-end COM path without going through MCP/stdio.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / "cycle_test.hwpx"

    def _write(hwp):
        hwp.add_doc()
        hwp.insert_text("안녕 세상")
        hwp.BreakPara()
        hwp.create_table(rows=2, cols=2, treat_as_char=False, header=True)
        hwp.save_as(str(target), format="HWPX")
        return str(target)

    wrote = session.call_sync(_write)
    print(f"[1/2] wrote: {wrote}")

    def _read(hwp):
        hwp.open(str(target), format="", arg="lock:true")
        # Use option="" to read the entire document. The default
        # "saveblock:true" only returns a selected block and yields None.
        return hwp.get_text_file("TEXT", "") or ""

    text = session.call_sync(_read)
    print(f"[2/2] read back {len(text)} chars")
    assert "안녕" in text, f"expected 안녕 in extracted text, got: {text[:200]!r}"
    print("OK - cycle smoke test passed")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
