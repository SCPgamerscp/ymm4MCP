"""Regression checks for hostile Unicode, oversized identities, and stable QA evidence."""
import unittest

import asset_qa


def snapshot(*, missing=None, unchecked=None, unknown=None, count=None):
    missing = [] if missing is None else missing
    unchecked = [] if unchecked is None else unchecked
    unknown = [] if unknown is None else unknown
    return {
        "success": True,
        "complete": not unknown and not unchecked,
        "checked_count": len(missing) if count is None else count,
        "missing": missing,
        "unknown_source_item_ids": unknown,
        "unchecked": unchecked,
    }


def source(path, item_id="native:1"):
    return {"path": path, "item_ids": [item_id]}


class AssetQaIdentityHardeningTests(unittest.TestCase):
    def test_rejects_unpaired_unicode_surrogates_in_ids(self):
        for bad_id in ("native:\ud800", "native:\udfff"):
            with self.subTest(id=repr(bad_id)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(unknown=[bad_id]))
            with self.subTest(grouped=repr(bad_id)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source("C:/media/a.mp4", bad_id)]))

    def test_rejects_unpaired_unicode_surrogates_in_paths(self):
        for path in ("C:/media/\ud800.mp4", "/media/\udfff.mp4"):
            for category in ("missing", "unchecked"):
                with self.subTest(path=repr(path), category=category), self.assertRaises(ValueError):
                    asset_qa.report(snapshot(**{category: [source(path)]},
                                             count=1 if category == "missing" else 0))

    def test_rejects_invisible_unicode_format_characters_in_ids(self):
        for char in ("\u200b", "\u2060", "\ufeff", "\u206a", "\u206f", "\u00ad"):
            with self.subTest(char=repr(char)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(unknown=["item" + char + "1"]))

    def test_rejects_invisible_unicode_format_characters_in_paths(self):
        for path in ("C:/media/a\u200bb.mp4", "/media/a\ufeffb.mp4"):
            with self.subTest(path=repr(path)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source(path)]))

    def test_allows_joined_emoji_and_unicode_identities(self):
        result = asset_qa.report(snapshot(missing=[source("/media/👩\u200d💻.mp4", "動画👩\u200d💻")]))
        self.assertEqual(result["issues"][0]["item_ids"], ["動画👩\u200d💻"])

    def test_rejects_oversized_item_id_but_accepts_boundary(self):
        with self.assertRaises(ValueError):
            asset_qa.report(snapshot(unknown=["a" * 257]))
        result = asset_qa.report(snapshot(unknown=["a" * 256]))
        self.assertEqual(len(result["issues"]), 1)

    def test_rejects_oversized_posix_and_windows_paths(self):
        for path in ("/" + "a" * 4096, "C:/" + "a" * 32767):
            with self.subTest(path_length=len(path)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source(path)]))

    def test_multibyte_path_limits_use_filesystem_code_units(self):
        # Python len() undercounts bytes on POSIX and UTF-16 surrogate pairs on Windows.
        for path in ("/" + "あ" * 1400, "C:/" + "😀" * 17000):
            with self.subTest(prefix=path[:3]), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source(path)]))

    def test_report_does_not_alias_missing_source_item_ids(self):
        original = source("C:/missing.mp4", "native:1")
        result = asset_qa.report(snapshot(missing=[original]))
        original["item_ids"].append("native:2")
        self.assertEqual(result["issues"][0]["item_ids"], ["native:1"])

    def test_report_does_not_alias_unknown_or_unchecked_item_ids(self):
        unknown = ["native:unknown"]
        unchecked = source("C:/unchecked.mp4", "native:unavailable")
        result = asset_qa.report(snapshot(unknown=unknown, unchecked=[unchecked], count=0))
        unknown.clear()
        unchecked["item_ids"].append("native:other")
        self.assertEqual(result["issues"][0]["item_ids"], ["native:unknown"])
        self.assertEqual(result["issues"][1]["item_ids"], ["native:unavailable"])


if __name__ == "__main__":
    unittest.main()
