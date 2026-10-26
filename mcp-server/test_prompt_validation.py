"""Regression tests for bounded production prompt arguments (no YMM4 needed)."""
import unittest
from prompt_validation import validate_prompt_arguments


class PromptValidationTests(unittest.TestCase):
    def test_valid_windows_paths_and_multiline_theme(self):
        for path in ("C:/素材/video.mp4", "C:\\素材\\video.mp4",
                     "\\\\server\\share\\video.mp4"):
            with self.subTest(path=path):
                args = {"theme": "冒頭\nテスト", "video_path": path,
                        "duration_seconds": "60"}
                self.assertIs(validate_prompt_arguments(args), args)

    def test_rejects_relative_drive_relative_urls_and_root_relative_paths(self):
        for path in ("clip.mp4", "../clip.mp4", "C:clip.mp4", "1:/clip.mp4", r"\\server", r"\\server\share", r"\\.\PhysicalDrive0", r"\\?\C:\Windows\System32", "/clip.mp4",
                     "https://example.com/video.mp4", "C:/clip\n.mp4"):
            with self.subTest(path=path), self.assertRaisesRegex(ValueError, "video_path"):
                validate_prompt_arguments({"theme": "実況", "video_path": path})

    def test_rejects_oversized_video_path(self):
        with self.assertRaisesRegex(ValueError, "video_path"):
            validate_prompt_arguments({"theme": "実況", "video_path": "C:/" + "a" * 4096})

    def test_rejects_oversized_or_invisible_theme(self):
        for theme in ("a" * 16385, "実況\x00", "実況\x7f", "\x80", "実況\x9f", " \t\n"):
            with self.subTest(theme=theme[:30]), self.assertRaisesRegex(ValueError, "theme"):
                validate_prompt_arguments({"theme": theme})

    def test_duration_bounds_and_exact_integer_format(self):
        for value in (1, "60", 86400, "86400"):
            with self.subTest(value=value):
                validate_prompt_arguments({"theme": "実況", "duration_seconds": value})
        for value in (0, -1, "86401", 86401, "1.0", 1.0, True, None, "-3"):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "duration_seconds"):
                validate_prompt_arguments({"theme": "実況", "duration_seconds": value})


if __name__ == "__main__":
    unittest.main()
