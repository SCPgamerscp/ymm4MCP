"""Asset QA must not trust malformed host source paths."""

import unittest

import asset_qa


class AssetQaAbsolutePathTests(unittest.TestCase):
    def _report(self, path, state="missing"):
        return {
            "success": True,
            "complete": state == "missing",
            "checked_count": 1 if state == "missing" else 0,
            "missing": [{"path": path, "item_ids": ["media:1"]}] if state == "missing" else [],
            "unknown_source_item_ids": [],
            "unchecked": [{"path": path, "item_ids": ["media:1"]}] if state == "unchecked" else [],
        }

    def test_rejects_relative_drive_relative_and_url_sources(self):
        for path in ("clip.mp4", "../clip.mp4", "C:clip.mp4",
                     "https://example.com/clip.mp4", "file:///C:/clip.mp4"):
            for state in ("missing", "unchecked"):
                with self.subTest(path=path, state=state), self.assertRaisesRegex(
                        ValueError, "valid path"):
                    asset_qa.report(self._report(path, state))

    def test_windows_drive_and_unc_sources_remain_valid(self):
        for path in ("C:/media/clip.mp4", "D:\\media\\clip.mp4",
                     "\\\\server\\share\\clip.mp4"):
            for state in ("missing", "unchecked"):
                with self.subTest(path=path, state=state):
                    result = asset_qa.report(self._report(path, state))
                    self.assertFalse(result["passed"])
                    self.assertEqual(len(result["issues"]), 1)


if __name__ == "__main__":
    unittest.main()
