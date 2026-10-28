"""Fail closed on ambiguous host asset identities and impossible source counts."""
import unittest

import asset_qa


class AssetQaIdentityAndLimitsTests(unittest.TestCase):
    def snapshot(self, *, missing=None, unknown=None, unchecked=None, count=1):
        missing = [] if missing is None else missing
        unknown = [] if unknown is None else unknown
        unchecked = [] if unchecked is None else unchecked
        return {
            "success": True,
            "complete": not unknown and not unchecked,
            "checked_count": count,
            "missing": missing,
            "unknown_source_item_ids": unknown,
            "unchecked": unchecked,
        }

    def source(self, path, item_id="native:1"):
        return {"path": path, "item_ids": [item_id]}

    def test_rejects_padded_unknown_and_grouped_item_ids(self):
        for item_id in (" native:1", "native:1 ", "\u00a0native:1", "native:1\u00a0"):
            with self.subTest(item_id=item_id):
                with self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(unknown=[item_id]))
                with self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(missing=[self.source("C:/missing.png", item_id)]))
                with self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(
                        unchecked=[self.source("C:/unchecked.png", item_id)], count=0))

    def test_preserves_unicode_and_internal_id_spaces(self):
        result = asset_qa.report(self.snapshot(
            missing=[self.source("C:/missing.png", "画像 その1")]))
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["item_ids"], ["画像 その1"])

    def test_rejects_checked_count_above_host_limit(self):
        for count in (101, 1000000000):
            with self.subTest(count=count), self.assertRaises(ValueError):
                asset_qa.report(self.snapshot(count=count))

    def test_rejects_combined_checked_and_unchecked_over_limit(self):
        with self.assertRaises(ValueError):
            asset_qa.report(self.snapshot(
                unchecked=[self.source("C:/unchecked.png")], count=100))
        # The exact producer boundary remains valid.
        result = asset_qa.report(self.snapshot(
            unchecked=[self.source("C:/unchecked.png")], count=99))
        self.assertFalse(result["passed"])

    def test_rejects_win32_trailing_dot_or_space_components(self):
        for path in ("C:/media./clip.mp4", "C:/media /clip.mp4",
                     "C:/media/clip.mp4.", "D:\\clips\\file.wav..",
                     "\\\\server\\share\\folder.\\clip.wav"):
            for state in ("missing", "unchecked"):
                with self.subTest(path=path, state=state), self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(
                        missing=[self.source(path)] if state == "missing" else [],
                        unchecked=[self.source(path)] if state == "unchecked" else [],
                        count=1 if state == "missing" else 0))

    def test_rejects_win32_alternate_data_stream_paths(self):
        for path in ("C:/media/clip.mp4:stream", "D:\\clip.wav::$DATA",
                     "\\\\server\\share\\clip.mp4:alt"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(self.snapshot(missing=[self.source(path)]))

    def test_rejects_windows_dot_segments_and_empty_components(self):
        # PR #230's shared export-path validator rejects traversal and
        # ambiguous Windows separators. Asset QA must honor the same rule
        # after all open PRs are composed; never weaken the path validator
        # merely to satisfy an obsolete compatibility expectation.
        for path in ("C:/media/../clip.mp4", "C:/media/./clip.mp4",
                     "C:/media//clip.mp4"):
            for state in ("missing", "unchecked"):
                with self.subTest(path=path, state=state), self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(
                        missing=[self.source(path)] if state == "missing" else [],
                        unchecked=[self.source(path)] if state == "unchecked" else [],
                        count=1 if state == "missing" else 0))

    def test_rejects_terminal_posix_dot_segments(self):
        # A final dot component aliases a directory, not a media file.
        # Exercise both report categories so neither bypasses validation.
        for path in ("/tmp/.", "/tmp/..", "/tmp/media/.", "/tmp/media/.."):
            for state in ("missing", "unchecked"):
                with self.subTest(path=path, state=state), self.assertRaises(ValueError):
                    asset_qa.report(self.snapshot(
                        missing=[self.source(path)] if state == "missing" else [],
                        unchecked=[self.source(path)] if state == "unchecked" else [],
                        count=1 if state == "missing" else 0))

    def test_preserves_valid_win32_and_posix_paths(self):
        for path in ("C:/My Assets/clip.mp4",
                     "\\\\server\\share\\voice clip.wav",
                     "/tmp/clip.with.dot.", "/tmp/clip:take.wav"):
            with self.subTest(path=path):
                result = asset_qa.report(self.snapshot(missing=[self.source(path)]))
                self.assertEqual(result["issues"][0]["code"], "ASSET_MISSING")


if __name__ == "__main__":
    unittest.main()
