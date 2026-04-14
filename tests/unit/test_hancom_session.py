"""Tests for HancomSession — mocks out pyhwpx so no real COM is needed.

We only verify the wiring: that ``call`` routes to the worker, that a
``HwpError`` propagates unchanged, and that an unknown exception resets
the cached instance.
"""

from __future__ import annotations

import asyncio
import sys
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.backend.errors import HwpError
from hwpx_hwp_mcp.backend.hancom_com import HancomSession


class _FakeHwp:
    def __init__(self):
        self.XHwpDocuments = MagicMock()
        self.XHwpDocuments.Count = 0

    def set_visible(self, *_):  # noqa: D401 - dummy
        pass

    def set_message_box_mode(self, *_):
        pass

    def quit(self, save=False):  # noqa: D401
        pass


@pytest.fixture
def patched_session(monkeypatch):
    fake = _FakeHwp()
    created = {"count": 0}

    def _make_session():
        sess = HancomSession()

        def _create_on_worker():
            created["count"] += 1
            return fake

        sess._create_on_worker = _create_on_worker  # type: ignore[method-assign]
        return sess

    sess = _make_session()
    try:
        yield sess, fake, created
    finally:
        asyncio.get_event_loop_policy().new_event_loop().run_until_complete(
            sess.shutdown()
        )


def test_call_sync_runs_on_single_thread(patched_session):
    sess, _fake, created = patched_session
    seen_threads: set[int] = set()

    def _task(hwp):
        seen_threads.add(threading.get_ident())
        return 42

    for _ in range(5):
        assert sess.call_sync(_task) == 42
    # The dedicated executor has exactly one worker.
    assert len(seen_threads) == 1
    # Session created exactly once.
    assert created["count"] == 1


def test_com_error_resets_instance(patched_session):
    sess, fake, created = patched_session

    def _raise(_hwp):
        raise RuntimeError("boom")

    # First call builds the instance, then errors; instance is reset.
    with pytest.raises(HwpError):
        sess.call_sync(_raise)
    assert sess._hwp is None  # noqa: SLF001
    # Next successful call recreates the instance.
    sess.call_sync(lambda _h: "ok")
    assert created["count"] == 2


def test_hwp_error_passes_through(patched_session):
    sess, _fake, _ = patched_session

    def _raise(_hwp):
        raise HwpError("specific message")

    with pytest.raises(HwpError, match="specific message"):
        sess.call_sync(_raise)
