"""Regression coverage for POSIX identity and Win32/Unicode asset QA ambiguity."""
import unittest

import asset_qa


def snapshot(*, missing=None, unchecked=None, unknown=None, count=None):
    missing = [] if missing is None else missing
    unchecked = [] if unchecked is None else unchecked
    unknown = [] if unknown is None else unknown
    if count is None:
        count = len(missing)
    return {
        "success": True,
        "complete": not unknown and not unchecked,
        "checked_count": count,
        "missing": missing,
        "unknown_source_item_ids": unknown,
        "unchecked": unchecked,
    }


def source(path, item_id):
    return {"path": path, "item_ids": [item_id]}


class AssetPathIdentityAndControlsTests(unittest.TestCase):
    def test_posix_case_sensitive_paths_remain_distinct(self):
        result = asset_qa.report(snapshot(missing=[
            source("/media/Scene.mp4", "video:1"),
            source("/media/scene.mp4", "video:2"),
        ]))
        self.assertEqual(len(result["issues"]), 2)
        self.assertFalse(result["passed"])

    def test_posix_backslash_is_literal_not_directory_separator(self):
        result = asset_qa.report(snapshot(missing=[
            source("/media/part\\take.wav", "audio:1"),
            source("/media/part/take.wav", "audio:2"),
        ]))
        self.assertEqual(len(result["issues"]), 2)

    def test_windows_case_and_separator_aliases_are_still_rejected(self):
        with self.assertRaises(ValueError):
            asset_qa.report(snapshot(missing=[
                source("C:/Media/Clip.wav", "audio:1"),
                source("c:\\media\\clip.wav", "audio:2"),
            ]))

    def test_rejects_embedded_c1_and_direction_controls_in_item_ids(self):
        for item_id in ("v\u0080x", "v\u009fx", "v\u202ex", "v\u2066x", "v\u061cx"):
            with self.subTest(item_id=repr(item_id)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(unknown=[item_id]))
            with self.subTest(source_id=repr(item_id)), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source("/media/a.mp4", item_id)]))

    def test_rejects_embedded_c1_and_direction_controls_in_paths(self):
        for path in ("/media/a\u0080b.mp4", "/media/a\u202eb.mp4",
                     "C:/media/a\u009fb.mp4", "C:/media/a\u2066b.mp4"):
            for state in ("missing", "unchecked"):
                with self.subTest(path=repr(path), state=state), self.assertRaises(ValueError):
                    asset_qa.report(snapshot(
                        missing=[source(path, "v1")] if state == "missing" else [],
                        unchecked=[source(path, "v1")] if state == "unchecked" else [],
                        count=1 if state == "missing" else 0))

    def test_rejects_unsafe_unc_server_and_share_components(self):
        for path in ("\\\\server.\\share\\clip.wav",
                     "\\\\server\\share.\\clip.wav",
                     "\\\\server\\share \\clip.wav",
                     "\\\\server\\CON\\clip.wav",
                     "\\\\server\\share:stream\\clip.wav"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source(path, "v1")]))

    def test_rejects_conin_and_conout_device_aliases(self):
        for path in ("C:/media/CONIN$.wav", "C:/CONOUT$.wav",
                     "\\\\nas\\share\\CONOUT$.mp4"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(snapshot(missing=[source(path, "v1")]))

    def test_valid_unicode_ids_posix_punctuation_and_unc_share_survive(self):
        for path in ("/media/日本語\\take.mp4", "\\\\nas\\media$\\日本語.mp4"):
            with self.subTest(path=path):
                result = asset_qa.report(snapshot(missing=[source(path, "動画 その1")]))
                self.assertEqual(result["issues"][0]["code"], "ASSET_MISSING")


if __name__ == "__main__":
    unittest.main()
