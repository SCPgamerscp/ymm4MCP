"""Regression: visually spoofed stable IDs must not authorize QA repair hints."""
import importlib.util
from pathlib import Path

import pytest

spec = importlib.util.spec_from_file_location("editing", Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py")
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)


@pytest.mark.parametrize("control", ["\u061c", "\u200e", "\u200f", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\u2066", "\u2067", "\u2068", "\u2069"])
def test_bidi_stable_id_never_used_for_repair(control):
    spoofed = "item" + control + "42"
    report = editing.validate_timeline(
        [{"item_id": spoofed, "frame": 0, "length": 20, "layer": 0}],
        duration=10,
    )
    assert not report["passed"]
    assert any(p["code"] == "INVALID_ITEM_ID" for p in report["issues"])
    assert all("suggested_fix" not in p for p in report["issues"])


def test_normal_unicode_stable_id_remains_usable():
    report = editing.validate_timeline(
        [{"item_id": "画像-42", "frame": 0, "length": 20, "layer": 0}],
        duration=10,
    )
    fixes = [p["suggested_fix"] for p in report["issues"] if "suggested_fix" in p]
    assert len(fixes) == 1
    assert fixes[0]["item_id"] == "画像-42"
