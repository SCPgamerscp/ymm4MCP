"""Reject ambiguous padded paths in host-provided asset QA snapshots (Issue #61)."""
import unittest

import asset_qa


class AssetQaPaddedPathsTests(unittest.TestCase):
    def _snapshot(self, state, path):
        return {
            "success": True,
            "complete": state == "missing",
            "checked_count": 1 if state == "missing" else 0,
            "missing": [{"path": path, "item_ids": ["asset:1"]}] if state == "missing" else [],
            "unchecked": [{"path": path, "item_ids": ["asset:1"]}] if state == "unchecked" else [],
            "unknown_source_item_ids": [],
        }

    def test_rejects_padded_missing_paths(self):
        for path in (" C:/assets/video.mp4", "C:/assets/video.mp4 ", "\\tC:/assets/video.mp4"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(self._snapshot("missing", path))

    def test_rejects_padded_unchecked_paths(self):
        for path in (" C:/assets/audio.wav", "C:/assets/audio.wav "):
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(self._snapshot("unchecked", path))

    def test_preserves_spaces_inside_valid_filenames(self):
        result = asset_qa.report(self._snapshot("missing", "C:/My Assets/voice clip.wav"))
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["code"], "ASSET_MISSING")


if __name__ == "__main__":
    unittest.main()
