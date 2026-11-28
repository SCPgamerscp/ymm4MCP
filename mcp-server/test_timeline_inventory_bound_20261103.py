"""Regression tests for bounded timeline QA input (#13/#62)."""
import unittest
from unittest.mock import patch

import editing
from editing import validate_timeline


class TimelineInventoryBoundTests(unittest.TestCase):
    def test_rejects_oversized_inventory_before_inspection(self):
        with self.assertRaisesRegex(ValueError, "at most 10000"):
            validate_timeline([None] * 10001)

    def test_boundary_is_accepted(self):
        result = validate_timeline([], include_gaps=False)
        self.assertTrue(result["passed"])
        result = validate_timeline([{"frame": 0, "length": 1, "layer": i}
                                    for i in range(10000)], include_gaps=False)
        self.assertEqual(result["item_count"], 10000)
        self.assertTrue(result["passed"])

    def test_invalid_input_types_fail_closed(self):
        for value in (None, {}, "items", ()):
            with self.subTest(value=value), self.assertRaises(ValueError):
                validate_timeline(value)

    def test_many_overlaps_reuse_layer_identity_validation(self):
        items = [{"item_id": f"native:{i}", "frame": i, "length": 2, "layer": 0}
                 for i in range(100)]
        with patch.object(editing, "_valid_item_id", wraps=editing._valid_item_id) as check_id:
            result = validate_timeline(items)
        self.assertEqual(result["issue_count"], 99)
        self.assertTrue(all(issue["suggested_fix"]["layers"] == [0] for issue in result["issues"]))
        self.assertLess(check_id.call_count, 10 * len(items))


if __name__ == "__main__":
    unittest.main()
