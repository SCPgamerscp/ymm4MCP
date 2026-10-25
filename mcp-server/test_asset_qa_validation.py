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

    def test_rejects_duplicate_unknown_ids(self):
        with self.assertRaises(ValueError):
            asset_qa.report({**self.good, "complete": False,
                             "unknown_source_item_ids": ["unknown", "unknown"]})

    def test_rejects_duplicate_ids_across_asset_states(self):
        missing = [{"path": "C:/missing.png", "item_ids": ["same"]}]
        for variant in (
            {"missing": [{"path": "C:/missing.png", "item_ids": ["same", "same"]}]},
            {"missing": missing, "unknown_source_item_ids": ["same"]},
            {"missing": missing,
             "unchecked": [{"path": "D:/unchecked.png", "item_ids": ["same"]}]},
            {"missing": missing + [{"path": "D:/missing.png", "item_ids": ["same"]}]},
        ):
            with self.subTest(variant=variant), self.assertRaises(ValueError):
                asset_qa.report({**self.good, "checked_count": 2,
                                 "complete": not (variant.get("unknown_source_item_ids") or
                                                  variant.get("unchecked")),
                                 **variant})

    def test_rejects_duplicate_canonical_paths(self):
        for variant in (
            {"missing": [{"path": "C:/a.png", "item_ids": ["a"]},
                         {"path": "c:\\\\a.png", "item_ids": ["b"]}]},
            {"missing": [{"path": "C:/a.png", "item_ids": ["a"]}],
             "unchecked": [{"path": "c:/a.png", "item_ids": ["b"]}]},
        ):
            with self.subTest(variant=variant), self.assertRaises(ValueError):
                asset_qa.report({**self.good, "checked_count": 2,
                                 "complete": not variant.get("unchecked"), **variant})

    def test_rejects_blank_and_control_character_ids(self):
        for bad_id in (" ", "  ", "\n", "a\x00b", "a\x7fb"):
            with self.subTest(bad_id=bad_id):
                with self.assertRaises(ValueError):
                    asset_qa.report({**self.good, "complete": False,
                                     "unknown_source_item_ids": [bad_id]})
                with self.assertRaises(ValueError):
                    asset_qa.report({**self.good,
                                     "missing": [{"path": "C:/missing", "item_ids": [bad_id]}]})

    def test_rejects_control_character_paths(self):
        for bad_path in ("C:/path\nother", "C:/a\x00b", "C:/a\x7fb"):
            with self.subTest(bad_path=bad_path), self.assertRaises(ValueError):
                asset_qa.report({**self.good,
                                 "missing": [{"path": bad_path, "item_ids": ["a"]}]})

    def test_accepts_distinct_grouped_paths_and_unicode_item_ids(self):
        result = asset_qa.report({
            "success": True, "complete": False, "checked_count": 2,
            "missing": [{"path": "C:/missing.png", "item_ids": ["画像1", "画像2"]}],
            "unknown_source_item_ids": ["未特定"],
            "unchecked": [{"path": "D:/unavailable.mp3", "item_ids": ["音声"]}],
        })
        self.assertFalse(result["passed"])
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["ASSET_MISSING", "ASSET_SOURCE_UNKNOWN", "ASSET_CHECK_UNAVAILABLE"])


if __name__ == "__main__":
    unittest.main()
