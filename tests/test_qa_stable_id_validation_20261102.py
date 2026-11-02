"""Regression coverage for stable IDs used by timeline QA repair suggestions."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from editing import validate_timeline


@pytest.mark.parametrize("bad_id", [
    " leading", "trailing ", "middle\x7fcontrol", "middle\x80control",
    "x" * 1025, "\ud800",
])
def test_invalid_stable_id_never_receives_automatic_trim_suggestion(bad_id):
    report = validate_timeline(
        [{"item_id": bad_id, "frame": 8, "length": 5, "layer": 0}],
        duration=10,
    )
    codes = {issue["code"] for issue in report["issues"]}
    assert "INVALID_ITEM_ID" in codes
    overrun = next(issue for issue in report["issues"] if issue["code"] == "EXCEEDS_DURATION")
    assert "suggested_fix" not in overrun
    assert not report["passed"]


def test_valid_stable_id_keeps_safe_trim_suggestion():
    report = validate_timeline(
        [{"item_id": "item-42", "frame": 8, "length": 5, "layer": 0}],
        duration=10,
    )
    overrun = next(issue for issue in report["issues"] if issue["code"] == "EXCEEDS_DURATION")
    assert overrun["suggested_fix"]["item_id"] == "item-42"
    assert overrun["suggested_fix"]["value"] == 2
