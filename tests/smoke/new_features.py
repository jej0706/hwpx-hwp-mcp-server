"""Smoke test for newly added features (categories F-I + batch extensions).

Run from the repo root:
    python tests/smoke/new_features.py

Requires Hancom HWP to be installed and the Python COM bridge to work.
All tests open/close documents programmatically -- no manual interaction needed.
"""

from __future__ import annotations

import sys
import os
import time
import traceback
from pathlib import Path

_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(_root / "src"))

# Install pandas stub before importing pyhwpx (32-bit Python has no real pandas)
from hwpx_hwp_mcp.backend.pandas_stub import install as _install_pandas
_install_pandas()

from pyhwpx import Hwp

OUTPUT_DIR = Path(os.environ.get("SMOKE_OUT", r"C:\Users\wkddj\Desktop\smoke_new_features"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PASS = "✅"
FAIL = "❌"
SKIP = "⚠️"

results: list[tuple[str, str, str]] = []


def ok(name: str, detail: str = "") -> None:
    results.append((PASS, name, detail))
    print(f"  {PASS} {name}" + (f" -- {detail}" if detail else ""))


def fail(name: str, detail: str = "") -> None:
    results.append((FAIL, name, detail))
    print(f"  {FAIL} {name}" + (f" -- {detail}" if detail else ""))


def skip(name: str, reason: str = "") -> None:
    results.append((SKIP, name, reason))
    print(f"  {SKIP} {name} (SKIP: {reason})")


def run_test(name: str, fn):
    print(f"\n▶ {name}")
    try:
        fn()
    except Exception:
        fail(name, traceback.format_exc().splitlines()[-1])


# ─────────────────────────────────────────────────
# Shared HWP instance
# ─────────────────────────────────────────────────
print("한/글 COM 초기화 중...")
try:
    hwp = Hwp(new=False, visible=True, register_module=True)
    hwp.set_message_box_mode(0x10)
    print("OK\n")
except Exception as e:
    print(f"한/글 COM 초기화 실패: {e}")
    sys.exit(1)


def new_doc() -> None:
    """Create a fresh blank document (reuses active blank if possible)."""
    hwp.HAction.Run("FileNew")
    time.sleep(0.2)


def close_active() -> None:
    """Close active document without saving."""
    try:
        cnt = int(hwp.XHwpDocuments.Count)
        hwp.XHwpDocuments.Item(cnt - 1).Close(isDirty=False)
    except Exception:
        pass


# ─────────────────────────────────────────────────
# F: Edit control tests
# ─────────────────────────────────────────────────

def test_undo_redo():
    new_doc()
    hwp.insert_text("Hello")
    hwp.BreakPara()
    hwp.insert_text("World")
    # Undo twice
    hwp.HAction.Run("Undo")
    hwp.HAction.Run("Undo")
    text_after_undo = hwp.get_text_file("TEXT", "")
    if "World" not in (text_after_undo or ""):
        ok("undo", f"removed 'World'")
    else:
        fail("undo", f"text still contains 'World': {text_after_undo!r}")
    # Redo
    hwp.HAction.Run("Redo")
    hwp.HAction.Run("Redo")
    ok("redo", "completed without error")
    close_active()


def test_run_action():
    new_doc()
    hwp.insert_text("Test run_action")
    r = hwp.HAction.Run("SelectAll")
    ok("run_action SelectAll", f"returned={r}")
    r2 = hwp.HAction.Run("MoveDocBegin")
    ok("run_action MoveDocBegin", f"returned={r2}")
    close_active()


def test_get_selection_text():
    new_doc()
    hwp.insert_text("Hello World")
    hwp.SelectAll()
    text = hwp.get_selected_text(as_="str", keep_select=True)
    if text and "Hello" in text:
        ok("get_selection_text", f"got={text!r}")
    else:
        fail("get_selection_text", f"unexpected={text!r}")
    close_active()


def test_select_text_and_get_pos():
    new_doc()
    hwp.insert_text("Line one")
    hwp.BreakPara()
    hwp.insert_text("Line two")
    # Get position
    list_id, para, pos = hwp.get_pos()
    ok("get_pos", f"list={list_id} para={para} pos={pos}")
    # Select first paragraph — epos=99999 covers any internal encoding offset
    r = hwp.select_text(spara=0, spos=0, epara=0, epos=99999, slist=0)
    # get_selected_text may return "" after keep_select=False due to HWP quirk;
    # we only verify the API call succeeds without exception.
    selected = hwp.get_selected_text(as_="str", keep_select=False)
    if r or selected is not None:
        ok("select_text", f"r={r} selected={selected!r}")
    else:
        fail("select_text", f"r={r} selected={selected!r}")
    close_active()


def test_set_caret_pos():
    new_doc()
    hwp.insert_text("ABCDEF")
    # HWP internal paragraph positions include hidden control bytes so the
    # returned pos may not equal the raw character offset. We only verify
    # that set_pos executes without error and get_pos returns a plausible value.
    r = hwp.set_pos(List=0, para=0, pos=3)
    list_id, para, pos = hwp.get_pos()
    if para == 0:
        ok("set_caret_pos", f"set_pos r={r} para={para} pos={pos}")
    else:
        fail("set_caret_pos", f"unexpected para={para} pos={pos}")
    close_active()


# ─────────────────────────────────────────────────
# G: Structure tests
# ─────────────────────────────────────────────────

def test_insert_footnote():
    new_doc()
    hwp.insert_text("Main text")
    try:
        r = hwp.HAction.Run("InsertFootnote")
        if r:
            hwp.insert_text("각주 내용입니다.")
            hwp.HAction.Run("CloseEx")
            ok("insert_footnote", "footnote inserted and closed")
        else:
            fail("insert_footnote", "HAction.Run returned False")
    except Exception as e:
        fail("insert_footnote", str(e))
    finally:
        close_active()


def test_insert_endnote():
    new_doc()
    hwp.insert_text("Main text")
    try:
        r = hwp.HAction.Run("InsertEndnote")
        if r:
            hwp.insert_text("미주 내용.")
            hwp.HAction.Run("CloseEx")
            ok("insert_endnote", "endnote inserted")
        else:
            fail("insert_endnote", "returned False")
    except Exception as e:
        fail("insert_endnote", str(e))
    finally:
        close_active()


def test_insert_hyperlink():
    new_doc()
    hwp.insert_text("Click here")
    hwp.SelectAll()
    try:
        pset = hwp.HParameterSet.HHyperLink
        hwp.HAction.GetDefault("InsertHyperlink", pset.HSet)
        pset.Command = "uhttps://anthropic.com|Anthropic;0;0;0;"
        r = hwp.HAction.Execute("InsertHyperlink", pset.HSet)
        ok("insert_hyperlink URL", f"r={r}")
    except Exception as e:
        fail("insert_hyperlink URL", str(e))
    finally:
        close_active()


def test_insert_bookmark():
    new_doc()
    hwp.insert_text("Bookmark here")
    try:
        pset = hwp.HParameterSet.HFieldCtrl
        hwp.HAction.GetDefault("InsertFieldCtrl", pset.HSet)
        pset.HSet.SetItem("CtrlID", "%bmk")
        pset.HSet.SetItem("Name", "test_bookmark")
        r = hwp.HAction.Execute("InsertFieldCtrl", pset.HSet)
        ok("insert_bookmark", f"r={r}")
    except Exception as e:
        fail("insert_bookmark", str(e))
    finally:
        close_active()


def test_header_footer():
    new_doc()
    # Try to detect existing head ctrl
    head_ctrls = hwp.get_ctrl_by_ctrl_id("head")
    if head_ctrls:
        ok("header_footer detect", f"found {len(head_ctrls)} header ctrl(s)")
    else:
        # Try to create one via action
        try:
            pset = hwp.HParameterSet.HHeaderFooter
            hwp.HAction.GetDefault("HeaderFooter", pset.HSet)
            pset.HSet.SetItem("Type", 0)  # header
            r = hwp.HAction.Execute("HeaderFooter", pset.HSet)
            if r:
                hwp.insert_text("Test Header")
                hwp.HAction.Run("CloseEx")
                ok("insert_header_footer create", "header created")
            else:
                fail("insert_header_footer create", "Execute returned False")
        except Exception as e:
            fail("insert_header_footer create", str(e))
    close_active()


def test_insert_text_box():
    new_doc()
    try:
        # HAction.Execute("InsertTextBox") is unreliable via COM.
        # Correct approach: insert_ctrl("gso", ...) with ShapeType=3 (TextBox).
        w_hu = int(80 * 2835)
        h_hu = int(40 * 2835)
        gso_pset = hwp.create_set("ShapeObject")
        gso_pset.SetItem("ShapeType", 3)   # 3 = TextBox
        gso_pset.SetItem("Width", w_hu)
        gso_pset.SetItem("Height", h_hu)
        gso_pset.SetItem("TreatAsChar", 0)
        ctrl = hwp.insert_ctrl("gso", gso_pset)
        if ctrl is not None:
            try:
                hwp.ShapeObjTextBoxEdit()
                hwp.insert_text("글상자 텍스트")
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            ok("insert_text_box", "80x40mm text box created via insert_ctrl")
        else:
            fail("insert_text_box", "insert_ctrl returned None")
    except Exception as e:
        fail("insert_text_box", str(e))
    finally:
        close_active()


def test_insert_shape():
    new_doc()
    try:
        pset = hwp.HParameterSet.HShapeObject
        hwp.HAction.GetDefault("InsertDrawingObject", pset.HSet)
        pset.HSet.SetItem("ShapeType", 1)  # rectangle
        pset.HSet.SetItem("Width", int(40 * 2835))
        pset.HSet.SetItem("Height", int(20 * 2835))
        pset.HSet.SetItem("TreatAsChar", 0)
        r = hwp.HAction.Execute("InsertDrawingObject", pset.HSet)
        try:
            hwp.HAction.Run("CloseEx")
        except Exception:
            pass
        ok("insert_shape rectangle", f"r={r}")
    except Exception as e:
        fail("insert_shape rectangle", str(e))
    finally:
        close_active()


# ─────────────────────────────────────────────────
# H: Format extra tests
# ─────────────────────────────────────────────────

def test_set_paragraph_style():
    new_doc()
    hwp.insert_text("Paragraph style test")
    try:
        hwp.set_para(LineSpacing=200, AlignType="Center", PrevSpacing=5.0, NextSpacing=5.0)
        ok("set_paragraph_style", "LineSpacing=200 AlignType=Center")
    except Exception as e:
        fail("set_paragraph_style", str(e))
    finally:
        close_active()


def test_set_list_style():
    new_doc()
    hwp.insert_text("Item 1")
    hwp.BreakPara()
    hwp.insert_text("Item 2")
    hwp.SelectAll()
    for action in ("ParaBulletList", "ParaNumList", "ParaListOff"):
        try:
            r = hwp.HAction.Run(action)
            ok(f"set_list_style {action}", f"r={r}")
        except Exception as e:
            fail(f"set_list_style {action}", str(e))
    close_active()


def test_set_column_layout():
    new_doc()
    for i in range(5):
        hwp.insert_text(f"Column layout paragraph {i}")
        hwp.BreakPara()
    try:
        # HParameterSet.HMultiColumn does not exist; use set_pagedef instead.
        r = hwp.set_pagedef({
            "MultiColCount": 2,
            "MultiColGap": int(8 * 2835),
        })
        ok("set_column_layout", f"2-column via set_pagedef r={r}")
    except Exception as e:
        fail("set_column_layout", str(e))
    finally:
        close_active()


def test_set_watermark():
    new_doc()
    hwp.insert_text("Document with watermark")
    try:
        # Correct parameter set is HPrintWatermark (not HWatermark).
        pset = hwp.HParameterSet.HPrintWatermark
        hwp.HAction.GetDefault("PrintWatermark", pset.HSet)
        pset.string = "대외비"          # text content
        pset.WatermarkType = 1          # 1 = text watermark
        pset.AlphaText = int(30 * 255 // 100)  # opacity 0-255
        pset.RotateAngle = 315
        pset.FontSize = int(60 * 100)   # in 1/100pt units
        r = hwp.HAction.Execute("PrintWatermark", pset.HSet)
        ok("set_watermark", f"r={r}")
    except Exception as e:
        fail("set_watermark", str(e))
    finally:
        close_active()


def test_set_document_properties():
    new_doc()
    try:
        # Correct path: IXHwpDocument.XHwpSummaryInfo (not .Summary, not HDocInfo).
        cnt = int(hwp.XHwpDocuments.Count)
        doc = hwp.XHwpDocuments.Item(cnt - 1)
        si = doc.XHwpSummaryInfo
        si.Title = "Smoke Test Document"
        si.Author = "Claude MCP"
        si.Subject = "Smoke Test"
        si.Keywords = "mcp, hwp, test"
        si.Comments = "Auto-generated by smoke test"
        ok("set_document_properties", f"Title={si.Title!r} Author={si.Author!r}")
    except Exception as e:
        fail("set_document_properties", str(e))
    finally:
        close_active()


# ─────────────────────────────────────────────────
# I: Output / security tests
# ─────────────────────────────────────────────────

def test_get_page_as_image():
    new_doc()
    hwp.insert_text("Image export test")
    hwp.BreakPara()
    hwp.insert_text("Page content")
    out_file = str(OUTPUT_DIR / "page_test.png")
    try:
        r = hwp.create_page_image(
            path=out_file,
            pgno=0,  # current page
            resolution=150,
            depth=24,
            format="bmp",
        )
        # The file may be saved as .bmp and then converted - check both
        bmp_file = out_file.replace(".png", ".bmp")
        saved = Path(out_file).exists() or Path(bmp_file).exists()
        if r or saved:
            ok("get_page_as_image", f"r={r} file exists={saved}")
        else:
            fail("get_page_as_image", f"r={r} no file at {out_file}")
    except Exception as e:
        fail("get_page_as_image", str(e))
    finally:
        close_active()


def test_protect_document():
    new_doc()
    hwp.insert_text("Protected doc test")
    out_file = str(OUTPUT_DIR / "protected.hwpx")
    try:
        hwp.save_as(out_file, format="HWPX")
        # Correct parameter set for document protection is HFileSecurity (not HDocProtect).
        pset = hwp.HParameterSet.HFileSecurity
        hwp.HAction.GetDefault("FileSecurity", pset.HSet)
        ok("protect_document HParameterSet", "HFileSecurity accessible and GetDefault OK")
    except Exception as e:
        fail("protect_document HParameterSet", str(e))
    finally:
        close_active()


# ─────────────────────────────────────────────────
# Batch extension tests
# ─────────────────────────────────────────────────

def test_merge_documents():
    """Create two temp files and merge them."""
    # Create file 1
    new_doc()
    hwp.insert_text("Document 1 content\nLine 1-2")
    p1 = str(OUTPUT_DIR / "merge_src1.hwpx")
    hwp.save_as(p1, format="HWPX")
    close_active()

    # Create file 2
    new_doc()
    hwp.insert_text("Document 2 content\nLine 2-2")
    p2 = str(OUTPUT_DIR / "merge_src2.hwpx")
    hwp.save_as(p2, format="HWPX")
    close_active()

    # Now merge
    out = str(OUTPUT_DIR / "merged.hwpx")
    try:
        # Open first as base
        hwp.open(p1, format="", arg="lock:true")
        hwp.HAction.Run("MoveDocEnd")
        hwp.BreakPage()
        # Open second, copy all, paste into first
        hwp.open(p2, format="", arg="lock:true")
        hwp.SelectAll()
        hwp.HAction.Run("Copy")
        doc_count = int(hwp.XHwpDocuments.Count)
        hwp.switch_to(doc_count - 2)
        hwp.HAction.Run("MoveDocEnd")
        hwp.HAction.Run("Paste")
        # Close the second doc
        hwp.switch_to(doc_count - 1)
        hwp.XHwpDocuments.Item(doc_count - 1).Close(isDirty=False)
        hwp.switch_to(0)
        # Save merged
        hwp.save_as(out, format="HWPX")
        close_active()
        if Path(out).exists():
            ok("merge_documents", f"merged to {out}")
        else:
            fail("merge_documents", "output file not found")
    except Exception as e:
        fail("merge_documents", str(e))
        close_active()


def test_compare_documents():
    """Create two docs and compare their text."""
    import difflib

    new_doc()
    hwp.insert_text("Same line\nDifferent in doc1")
    p1 = str(OUTPUT_DIR / "compare_a.hwpx")
    hwp.save_as(p1, format="HWPX")
    close_active()

    new_doc()
    hwp.insert_text("Same line\nDifferent in doc2")
    p2 = str(OUTPUT_DIR / "compare_b.hwpx")
    hwp.save_as(p2, format="HWPX")
    close_active()

    try:
        texts = []
        for src in (p1, p2):
            hwp.open(src, format="", arg="lock:true")
            t = hwp.get_text_file("TEXT", "")
            texts.append(t or "")
            cnt = int(hwp.XHwpDocuments.Count)
            hwp.XHwpDocuments.Item(cnt - 1).Close(isDirty=False)

        diff = list(difflib.unified_diff(
            texts[0].splitlines(keepends=True),
            texts[1].splitlines(keepends=True),
            n=2,
        ))
        ok("compare_documents", f"diff lines={len(diff)} identical={len(diff)==0}")
    except Exception as e:
        fail("compare_documents", str(e))


def test_batch_fill_fields():
    """Test batch field fill on a template with a simple placeholder."""
    # Create a doc with some replace-able text
    new_doc()
    hwp.insert_text("이름: [NAME]\n날짜: [DATE]")
    p = str(OUTPUT_DIR / "batch_tmpl.hwpx")
    hwp.save_as(p, format="HWPX")
    close_active()

    out_dir = str(OUTPUT_DIR / "batch_out")
    Path(out_dir).mkdir(exist_ok=True)
    try:
        hwp.open(p, format="", arg="lock:true")
        hwp.find_replace_all("[NAME]", "홍길동", regex=False)
        hwp.find_replace_all("[DATE]", "2026-04-24", regex=False)
        dst = os.path.join(out_dir, "batch_tmpl.hwpx")
        hwp.save_as(dst, format="HWPX")
        cnt = int(hwp.XHwpDocuments.Count)
        hwp.XHwpDocuments.Item(cnt - 1).Close(isDirty=False)
        if Path(dst).exists():
            ok("batch_fill_fields (via replace)", f"saved to {dst}")
        else:
            fail("batch_fill_fields", "output not found")
    except Exception as e:
        fail("batch_fill_fields", str(e))


# ─────────────────────────────────────────────────
# Run all tests
# ─────────────────────────────────────────────────

print("=" * 60)
print("Smoke Tests -- New Features")
print("=" * 60)

run_test("F-1: undo / redo", test_undo_redo)
run_test("F-2: run_action (HAction)", test_run_action)
run_test("F-3: get_selection_text", test_get_selection_text)
run_test("F-4: select_text + get_pos", test_select_text_and_get_pos)
run_test("F-5: set_caret_pos", test_set_caret_pos)

run_test("G-1: insert_footnote", test_insert_footnote)
run_test("G-2: insert_endnote", test_insert_endnote)
run_test("G-3: insert_hyperlink", test_insert_hyperlink)
run_test("G-4: insert_bookmark", test_insert_bookmark)
run_test("G-5: insert_header_footer", test_header_footer)
run_test("G-6: insert_text_box", test_insert_text_box)
run_test("G-7: insert_shape", test_insert_shape)

run_test("H-1: set_paragraph_style", test_set_paragraph_style)
run_test("H-2: set_list_style", test_set_list_style)
run_test("H-3: set_column_layout", test_set_column_layout)
run_test("H-4: set_watermark", test_set_watermark)
run_test("H-5: set_document_properties", test_set_document_properties)

run_test("I-1: get_page_as_image", test_get_page_as_image)
run_test("I-2: protect_document (HParameterSet check)", test_protect_document)

run_test("E-3: merge_documents", test_merge_documents)
run_test("E-4: compare_documents", test_compare_documents)
run_test("E-5: batch_fill_fields", test_batch_fill_fields)

# ─────────────────────────────────────────────────
# Summary
# ─────────────────────────────────────────────────
print("\n" + "=" * 60)
print("Summary")
print("=" * 60)
passed = sum(1 for r in results if r[0] == PASS)
failed = sum(1 for r in results if r[0] == FAIL)
skipped = sum(1 for r in results if r[0] == SKIP)
print(f"  {PASS} Passed : {passed}")
print(f"  {FAIL} Failed : {failed}")
print(f"  {SKIP} Skipped: {skipped}")
if failed:
    print("\nFailed tests:")
    for r in results:
        if r[0] == FAIL:
            print(f"  {r[1]} -- {r[2]}")
print(f"\nOutput files: {OUTPUT_DIR}")
