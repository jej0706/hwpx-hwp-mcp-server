"""Manual smoke test: save a document as hwp, hwpx, pdf, docx.

This is the fastest way to verify the extension-dispatch path through
``pyhwpx.Hwp.save_as`` on your machine.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session


def main() -> None:
    out_dir = Path(__file__).resolve().parent / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)

    targets = {
        "hwp": out_dir / "sample.hwp",
        "hwpx": out_dir / "sample.hwpx",
        "pdf": out_dir / "sample.pdf",
        "docx": out_dir / "sample.docx",
    }

    def _write(hwp):
        hwp.add_doc()
        hwp.insert_text("Format roundtrip test 샘플")
        hwp.BreakPara()
        results: dict[str, bool] = {}
        for name, path in targets.items():
            fmt = "HWPX" if name == "hwpx" else ("PDF" if name == "pdf" else ("OOXML" if name == "docx" else "HWP"))
            ok = bool(hwp.save_as(str(path), format=fmt))
            results[name] = ok
        return results

    results = session.call_sync(_write)
    for name, path in targets.items():
        ok = results.get(name, False)
        size = path.stat().st_size if path.exists() else 0
        print(f"{name:>5}: ok={ok} size={size} path={path}")
        assert ok and size > 0, f"{name} export failed"
    print("OK - all formats exported")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
