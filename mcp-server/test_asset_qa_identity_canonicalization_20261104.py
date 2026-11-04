"""Regression tests for asset QA item identity canonicalization."""

import unittest

from asset_qa import report


def snapshot(item_id):
    return {
        "success": True,
        "complete": True,
        "checked_count": 0,
        "missing": [],
        "unchecked": [],
        "unknown_source_item_ids": [item_id],
    }


class AssetQaIdentityTests(unittest.TestCase):
    def test_rejects_zero_width_format_controls(self):
        for item_id in ("scene\u200c1", "scene\u200d1", "scene\u2060one"):
            with self.subTest(item_id=repr(item_id)):
                with self.assertRaises(ValueError):
                    report(snapshot(item_id))

    def test_rejects_non_nfc_alias(self):
        with self.assertRaises(ValueError):
            report(snapshot("Cafe\u0301"))

    def test_accepts_canonical_international_identifier(self):
        result = report(snapshot("カフェ_Café"))
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["item_ids"], ["カフェ_Café"])

    def test_rejects_unsafe_id_in_missing_source(self):
        source = {"path": "/media/clip.mp4", "item_ids": ["clip\u200d1"]}
        payload = {
            "success": True, "complete": True, "checked_count": 1,
            "missing": [source], "unknown_source_item_ids": [], "unchecked": [],
        }
        with self.assertRaises(ValueError):
            report(payload)


if __name__ == "__main__":
    unittest.main()
