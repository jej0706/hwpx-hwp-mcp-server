"""Smoke test: force_new=True still reuses a blank active document.

This is a regression guard for a user-reported bug where saying
"새 한글 창을 열고 표 넣어줘" caused Claude to call
``create_new_document(force_new=True)``. Before the fix, that spawned a
빈 문서 2 on top of an empty 빈 문서 1, and on Hancom 2024 installs
without tab UI the original window appeared to "vanish into the
background" while the new tab took over. The desired behavior is: if
the user's active document is already blank, reuse it silently
regardless of ``force_new``.

Scenarios covered
-----------------

1. Pre-seed a single blank document. Call the equivalent of
   ``create_new_document(force_new=True)``. Assert that:
     - the document count does NOT increase, and
     - the returned doc_id is the same as the pre-existing active doc.

2. Force the active document to look non-blank (insert text so Modified
   becomes True). Call the force_new path again. Assert that:
     - the document count DOES increase by 1, and
     - the returned doc_id is the newly-added one.

We deliberately skip any file system writes - Modified=True via
``insert_text`` is enough to flip ``_is_blank_doc`` to False.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.hancom_com import session
from hwpx_hwp_mcp.tools.session import (
    _count,
    _doc_ref_from_active,
    _get_active_doc_id,
    _is_blank_doc,
    _require_doc,
)


def _simulate_create_new_document_force_new(hwp):
    """Mirror the body of create_new_document(force_new=True)."""
    active_idx = _get_active_doc_id(hwp)
    if active_idx >= 0 and _is_blank_doc(hwp, active_idx):
        _require_doc(hwp, active_idx)
        return _doc_ref_from_active(hwp, active_idx).doc_id, "reused"
    hwp.add_tab()
    doc_id = max(0, _count(hwp) - 1)
    _require_doc(hwp, doc_id)
    return doc_id, "created"


def main() -> None:
    # --- Scenario 1: blank active doc -> force_new is ignored ------------

    def seed_blank(hwp):
        if _count(hwp) == 0:
            hwp.add_doc()
        # Make sure the active doc is a pristine blank one. If user had
        # content we don't touch it - but for a CI-ish smoke we assume
        # scenario 1 runs against a fresh state.
        return _count(hwp)

    count_before = session.call_sync(seed_blank)
    print(f"[seed] docs before blank-reuse test: {count_before}")
    assert count_before >= 1, "failed to seed at least one blank document"

    active_before = session.call_sync(_get_active_doc_id)
    blank_before = session.call_sync(lambda h: _is_blank_doc(h, active_before))
    print(f"[seed] active_idx={active_before} is_blank={blank_before}")
    assert blank_before, (
        "scenario 1 needs a blank active document - "
        "close any non-empty docs in 한글 and re-run."
    )

    returned_id, action = session.call_sync(_simulate_create_new_document_force_new)
    count_after = session.call_sync(_count)
    print(
        f"[blank-reuse] returned_id={returned_id} action={action} "
        f"count_after={count_after}"
    )
    assert action == "reused", f"expected reuse, got action={action}"
    assert count_after == count_before, (
        f"count changed from {count_before} to {count_after} - "
        "blank-reuse guard did NOT block the new tab."
    )
    assert returned_id == active_before, (
        f"returned doc_id {returned_id} != active {active_before}"
    )
    print("[1/2] blank-reuse OK - force_new ignored on an empty doc")

    # --- Scenario 2: non-blank active doc -> force_new creates new tab --

    def dirty_the_active(hwp):
        # insert_text via HAction flips Modified=True.
        idx = _get_active_doc_id(hwp)
        _require_doc(hwp, idx)
        hwp.insert_text("not blank any more")
        return idx, _is_blank_doc(hwp, idx)

    idx, blank_after_edit = session.call_sync(dirty_the_active)
    print(f"[dirty] active_idx={idx} is_blank={blank_after_edit}")
    assert not blank_after_edit, (
        "insert_text did not flip Modified - smoke test cannot continue."
    )

    count_before2 = session.call_sync(_count)
    returned_id2, action2 = session.call_sync(_simulate_create_new_document_force_new)
    count_after2 = session.call_sync(_count)
    print(
        f"[dirty-force] returned_id={returned_id2} action={action2} "
        f"count: {count_before2} -> {count_after2}"
    )
    assert action2 == "created", f"expected created, got {action2}"
    assert count_after2 == count_before2 + 1, (
        f"expected count to grow by 1, got {count_before2} -> {count_after2}"
    )
    assert returned_id2 != idx, (
        f"returned doc_id {returned_id2} should differ from dirty doc {idx}"
    )
    print("[2/2] dirty-force OK - new tab created when active has content")

    print("OK - blank-reuse smoke test passed")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
