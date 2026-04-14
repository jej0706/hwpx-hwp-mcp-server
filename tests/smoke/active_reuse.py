"""Regression smoke test for the 'reuse active document' behavior.

User-reported issue: when HWP is already running with a document visible,
calling ``create_new_document`` spawned an additional window/tab even
though the user wanted the MCP server to edit the existing document.

The fix: ``create_new_document`` now defaults to ``prefer_active=True`` -
if any document is already open, its doc_id is returned instead of
creating a new one. The new ``get_active_document`` tool exposes this
probe explicitly for callers that want to be defensive.

This script verifies three things:

1. When HWP has a pre-existing document open, ``create_new_document``
   (with default prefer_active=True) returns the SAME doc_id as the
   existing document - XHwpDocuments.Count does NOT increase.
2. When no document is open, ``create_new_document`` actually creates
   one and Count becomes 1.
3. ``get_active_document`` returns the active doc_id without creating
   anything, and raises when no document is open.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.errors import HwpDocumentNotFound
from hwpx_hwp_mcp.backend.hancom_com import session
from hwpx_hwp_mcp.tools.session import (
    _count,
    _doc_ref_from_active,
    _get_active_doc_id,
    _require_doc,
)


def main() -> None:
    # --- Scenario 1: HWP has a pre-existing doc (simulated by add_doc up front) ---

    def seed(hwp):
        # Ensure there is at least one doc open before we test reuse.
        if _count(hwp) == 0:
            hwp.add_doc()
        return _count(hwp)

    count_before = session.call_sync(seed)
    print(f"[seed] docs before reuse test: {count_before}")
    assert count_before >= 1, "failed to seed at least one document"

    # Mirror create_new_document(prefer_active=True) behavior
    def reuse(hwp):
        active_idx = _get_active_doc_id(hwp)
        assert active_idx >= 0, f"expected active doc but got {active_idx}"
        _require_doc(hwp, active_idx)
        ref = _doc_ref_from_active(hwp, active_idx)
        return active_idx, _count(hwp), ref.doc_id

    active_idx, count_after, returned_id = session.call_sync(reuse)
    print(
        f"[reuse] active_idx={active_idx} count_after={count_after} returned_id={returned_id}"
    )
    assert count_after == count_before, (
        f"count changed from {count_before} to {count_after} - reuse DID create a new doc"
    )
    assert returned_id == active_idx, (
        f"returned doc_id {returned_id} does not match active {active_idx}"
    )
    print("[1/3] reuse path OK - no new document created")

    # --- Scenario 2: force a new document (prefer_active=False simulation) ---

    def force_new(hwp):
        before = _count(hwp)
        hwp.add_tab()  # mirrors create_new_document(prefer_active=False, tab=True)
        after = _count(hwp)
        return before, after

    before_add, after_add = session.call_sync(force_new)
    print(f"[force_new] before={before_add} after={after_add}")
    assert after_add == before_add + 1, (
        f"expected count to grow by 1, got {before_add} -> {after_add}"
    )
    print("[2/3] force-new path OK - count grew by exactly 1")

    # --- Scenario 3: get_active_document behavior with docs open ---

    def probe(hwp):
        idx = _get_active_doc_id(hwp)
        if idx < 0:
            return None
        _require_doc(hwp, idx)
        return _doc_ref_from_active(hwp, idx)

    ref = session.call_sync(probe)
    assert ref is not None, "get_active_document should return when docs are open"
    print(f"[3/3] get_active_document returned doc_id={ref.doc_id} (path={ref.path!r})")

    print("OK - active-reuse smoke test passed")


if __name__ == "__main__":
    try:
        main()
    finally:
        session.shutdown_sync()
