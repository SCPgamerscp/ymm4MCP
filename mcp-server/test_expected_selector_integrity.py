"""Expected timeline QA selectors must not match invalid or reused items."""
import unittest

from editing import validate_timeline


ITEMS = [
    {"item_id": "native:one", "revision": "r1", "frame": 0, "length": 10, "layer": 1, "type": "TextItem", "text": ""},
    {"item_id": "native:two", "revision": "r2", "frame": 20, "length": 10, "layer": 2, "type": "TextItem", "text": "hello"},
]


class ExpectedSelectorIntegrityTests(unittest.TestCase):
    def test_boolean_and_missing_value_selectors_cannot_match(self):
        for selector in ({"frame": False}, {"frame": True}, {"layer": True},
                         {"length": True}, {"frame": None}, {"layer": None},
                         {"type": None}, {"revision": None}, {"text": None}):
            with self.subTest(selector=selector), self.assertRaises(ValueError):
                validate_timeline(ITEMS, expected=[selector], include_gaps=False)

    def test_out_of_range_and_non_integer_frames_are_rejected(self):
        for selector in ({"frame": -1}, {"length": 0}, {"layer": 1.0},
                         {"frame": "0"}, {"length": 2147483648}):
            with self.subTest(selector=selector), self.assertRaises(ValueError):
                validate_timeline(ITEMS, expected=[selector], include_gaps=False)

    def test_null_or_blank_legacy_and_stable_ids_rejected(self):
        for selector in ({"id": None}, {"id": ""}, {"item_id": "  "},
                         {"revision": ""}, {"type": ""}):
            with self.subTest(selector=selector), self.assertRaises(ValueError):
                validate_timeline(ITEMS, expected=[selector], include_gaps=False)

    def test_duplicate_expected_selectors_cannot_reuse_one_item(self):
        report = validate_timeline(ITEMS, expected=[
            {"item_id": "native:one"}, {"frame": 0, "layer": 1}
        ], include_gaps=False)
        self.assertFalse(report["passed"])
        self.assertEqual([p["code"] for p in report["issues"]], ["EXPECTED_REUSED"])
        self.assertEqual(report["issues"][0]["item_ids"], ["native:one"])

    def test_legacy_alias_cannot_reuse_item(self):
        report = validate_timeline(ITEMS, expected=[
            {"id": "native:one"}, {"item_id": "native:one"}
        ], include_gaps=False)
        self.assertEqual(report["issues"][0]["code"], "EXPECTED_REUSED")

    def test_distinct_expected_items_and_empty_caption_still_pass(self):
        report = validate_timeline(ITEMS, expected=[
            {"id": "native:one", "text": ""},
            {"item_id": "native:two", "frame": 20, "layer": 2}
        ], include_gaps=False)
        self.assertTrue(report["passed"])
        self.assertEqual(report["issues"], [])

    def test_ambiguous_and_missing_selectors_keep_existing_codes(self):
        report = validate_timeline(ITEMS, expected=[
            {"type": "TextItem"}, {"item_id": "missing"}
        ], include_gaps=False)
        self.assertEqual([p["code"] for p in report["issues"]],
                         ["EXPECTED_AMBIGUOUS", "EXPECTED_NOT_FOUND"])


if __name__ == "__main__":
    unittest.main()
