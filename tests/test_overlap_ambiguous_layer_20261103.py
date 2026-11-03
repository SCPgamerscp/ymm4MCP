"""Regression tests for safe overlap repair suggestions."""
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location(
    "editing", Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py"
)
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)


def test_overlap_with_unique_ids_has_repair_hint():
    items = [
        {"item_id": "first", "frame": 0, "length": 20, "layer": 1},
        {"item_id": "second", "frame": 10, "length": 20, "layer": 1},
    ]
    issues = editing.validate_timeline(items)["issues"]
    overlaps = [issue for issue in issues if issue["code"] == "OVERLAP"]
    assert len(overlaps) == 1
    assert overlaps[0]["suggested_fix"]["sub_action"] == "resolve_overlaps"


def test_overlap_without_ids_has_no_mutation_hint():
    items = [
        {"frame": 0, "length": 20, "layer": 1},
        {"frame": 10, "length": 20, "layer": 1},
    ]
    issues = editing.validate_timeline(items)["issues"]
    overlaps = [issue for issue in issues if issue["code"] == "OVERLAP"]
    assert len(overlaps) == 1
    assert "suggested_fix" not in overlaps[0]


def test_unrelated_ambiguous_item_on_same_layer_blocks_layer_wide_repair():
    items = [
        {"item_id": "first", "frame": 0, "length": 20, "layer": 1},
        {"item_id": "second", "frame": 10, "length": 20, "layer": 1},
        {"item_id": "first", "frame": 40, "length": 10, "layer": 1},
    ]
    issues = editing.validate_timeline(items)["issues"]
    overlaps = [issue for issue in issues if issue["code"] == "OVERLAP"]
    assert len(overlaps) == 1
    assert "suggested_fix" not in overlaps[0]
