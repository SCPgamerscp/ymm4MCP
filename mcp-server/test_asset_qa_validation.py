"""Malformed host asset reports must not pass final QA."""
import unittest
import asset_qa


class AssetQaValidationTests(unittest.TestCase):
    def setUp(self):
        self.good = {"success": True, "complete": True, "checked_count": 1,
                     "missing": [], "unknown_source_item_ids": [], "unchecked": []}

    def test_valid_clean_and_missing_reports(self):
        self.assertTrue(asset_qa.report(self.good)["passed"])
        missing = {**self.good, "missing": [{"path": "C:/gone.png", "item_ids": ["img"]}]}
        self.assertEqual(asset_qa.report(missing)["issues"][0]["code"], "ASSET_MISSING")

    def test_rejects_malformed_source_entries(self):
        for source in (None, {}, {"path": "", "item_ids": ["a"]},
                       {"path": "C:/a", "item_ids": []},
                       {"path": "C:/a", "item_ids": [42]}):
            with self.subTest(source=source), self.assertRaises(ValueError):
                asset_qa.report({**self.good, "missing": [source]})

    def test_rejects_invalid_unknown_ids(self):
        with self.assertRaises(ValueError):
            asset_qa.report({**self.good, "complete": False,
                             "unknown_source_item_ids": [None]})

    def test_rejects_inconsistent_completeness(self):
        for complete in (True, None, "true"):
            with self.subTest(complete=complete), self.assertRaises(ValueError):
                asset_qa.report({**self.good, "complete": complete,
                                 "unchecked": [{"path": "C:/a", "item_ids": ["x"]}]})

    def test_rejects_invalid_checked_counts(self):
        for count in (None, True, -1, "1"):
            with self.subTest(count=count), self.assertRaises(ValueError):
                asset_qa.report({**self.good, "checked_count": count})
        with self.assertRaises(ValueError):
            asset_qa.report({**self.good, "checked_count": 0,
                             "missing": [{"path": "C:/gone", "item_ids": ["x"]}]})


if __name__ == "__main__":
    unittest.main()
