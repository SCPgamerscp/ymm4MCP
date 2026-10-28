"""Regression coverage for asset QA path identity and media-file requirements."""
import unittest

import asset_qa


class AssetQaWindowsPathIntegrityTests(unittest.TestCase):
    def snapshot(self, path):
        return {
            "success": True,
            "complete": True,
            "checked_count": 1,
            "missing": [{"path": path, "item_ids": ["asset:1"]}],
            "unknown_source_item_ids": [],
            "unchecked": [],
        }

    def assert_rejected(self, paths):
        for path in paths:
            with self.subTest(path=path), self.assertRaises(ValueError):
                asset_qa.report(self.snapshot(path))

    def test_rejects_windows_device_namespaces(self):
        self.assert_rejected((
            "\\\\?\\C:\\media\\clip.mp4",
            "\\\\.\\C:\\media\\clip.mp4",
            "//?/C:/media/clip.mp4",
            "//./C:/media/clip.mp4",
        ))

    def test_rejects_reserved_dos_device_basename_with_extensions(self):
        self.assert_rejected((
            "C:/media/CON.mp4",
            "C:/media/PRN",
            "C:/media/AUX.wav",
            "C:/media/NUL.mp4",
            "C:/media/COM1.wav",
            "C:/media/LPT9.txt",
            "C:/media/COM¹.mp4",
            "C:/media/LPT³.wav",
            "\\\\server\\share\\NUL.mp4",
        ))

    def test_rejects_invalid_windows_filename_characters(self):
        self.assert_rejected((
            "C:/media/clip?.mp4",
            "C:/media/clip*.mp4",
            "C:/media/clip|v2.mp4",
            "C:/media/clip<draft>.mp4",
            'C:/media/clip"v2.mp4',
        ))

    def test_rejects_unc_share_root_and_drive_root_as_media(self):
        self.assert_rejected((
            "\\\\server\\share",
            "C:/",
            "D:\\",
        ))

    def test_rejects_posix_directory_terminators(self):
        self.assert_rejected((
            "/",
            "/tmp/clip/",
        ))

    def test_preserves_valid_names_and_posix_punctuation(self):
        for path in (
            "C:/My Assets/clip.mp4",
            "C:/media/conductor.mp4",
            "C:/media/COM10.wav",
            "\\\\server\\share\\voice clip.wav",
            "/tmp/clip?draft.mp4",
            "/tmp/clip:take.wav",
            "/tmp/clip\\",
        ):
            with self.subTest(path=path):
                result = asset_qa.report(self.snapshot(path))
                self.assertFalse(result["passed"])
                self.assertEqual(result["issues"][0]["code"], "ASSET_MISSING")


if __name__ == "__main__":
    unittest.main()
