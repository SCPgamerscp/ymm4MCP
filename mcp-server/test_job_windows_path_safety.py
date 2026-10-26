"""Export and project paths must not target Windows devices or ambiguous files."""
import unittest

import jobs


class WindowsJobPathSafetyTests(unittest.TestCase):
    def test_valid_drive_unc_and_posix_paths(self):
        for path in ("C:/Videos/final.mp4", r"C:\Videos\final.mp4",
                     r"\\nas\share\final.mp4", r"\\nas\share\dir\final.mp4",
                     "/tmp/final.mp4"):
            with self.subTest(path=path):
                self.assertTrue(jobs.is_absolute_media_path(path))
                self.assertEqual(jobs.validate_export_request({"path": path})["format"], "mp4")
        self.assertEqual(jobs.validate_project_path({"path": r"\\nas\share\edit.ymmp"}),
                         r"\\nas\share\edit.ymmp")

    def test_rejects_incomplete_unc_and_windows_device_namespaces(self):
        for path in (r"\\nas", r"\\nas\share", r"\\.\PhysicalDrive0",
                     r"\\.\NUL\final.mp4", r"\\?\C:\Videos\final.mp4",
                     r"\\?\GLOBALROOT\Device\HarddiskVolume1\final.mp4",
                     "\\\\nas\\share\\", "//nas/share"):
            with self.subTest(path=path):
                self.assertFalse(jobs.is_absolute_media_path(path))

    def test_rejects_unsafe_windows_file_components(self):
        for path in ("C:/Videos/../final.mp4", "C:/Videos/./final.mp4",
                     "C:/Videos//final.mp4", "C:/Videos/trailing./final.mp4",
                     "C:/Videos/trailing /final.mp4", "C:/Videos:stream/final.mp4",
                     "C:/Videos/file?.mp4", "C:/Videos/file|name.mp4",
                     "C:/Videos/CON.mp4", "C:/Videos/NUL.mp4",
                     "C:/Videos/LPT1.mp4", "C:/Videos/COM9.mp4",
                     r"\\nas\share\AUX.mp4"):
            with self.subTest(path=path):
                self.assertFalse(jobs.is_absolute_media_path(path))
                with self.assertRaises(ValueError):
                    jobs.validate_export_request({"path": path})

    def test_project_paths_share_same_file_safety_boundary(self):
        for path in ("C:/projects/NUL.ymmp", "C:/projects/../edit.ymmp",
                     r"\\.\NUL\edit.ymmp", r"\\nas\share\CON.ymmp"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                jobs.validate_project_path({"path": path})


if __name__ == "__main__":
    unittest.main()
