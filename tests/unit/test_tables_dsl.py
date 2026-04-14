"""Unit tests for the fill_by_path DSL parser."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp.utils.tables import TablePath, parse_path


class TestParsePath:
    def test_single_direction(self):
        p = parse_path("이름: > right")
        assert p.label == "이름:"
        assert p.directions == ("right",)

    def test_multiple_directions(self):
        p = parse_path("합계 > down > down")
        assert p.label == "합계"
        assert p.directions == ("down", "down")

    def test_mixed_directions(self):
        p = parse_path("Header > right > down")
        assert p.directions == ("right", "down")

    def test_directions_are_lowercased(self):
        p = parse_path("X > RIGHT > Down")
        assert p.directions == ("right", "down")

    def test_label_may_contain_spaces(self):
        p = parse_path("총 합계 > down")
        assert p.label == "총 합계"

    def test_rejects_bare_label(self):
        with pytest.raises(ValueError):
            parse_path("이름")

    def test_rejects_empty(self):
        with pytest.raises(ValueError):
            parse_path("")

    def test_rejects_unknown_direction(self):
        with pytest.raises(ValueError):
            parse_path("X > sideways")

    def test_apply_moves_correctly(self):
        p = TablePath(label="L", directions=("right", "right", "down"))
        assert p.apply(1, 1) == (2, 3)

    def test_apply_handles_negatives(self):
        p = TablePath(label="L", directions=("left", "up"))
        assert p.apply(5, 5) == (4, 4)
