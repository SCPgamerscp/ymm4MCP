"""Regression tests for invisible format controls in prompt output paths."""

import unittest

from prompt_validation import validate_prompt_arguments


class PromptPathFormatControlTests(unittest.TestCase):
    def test_invisible_format_controls_are_rejected(self):
        for char in ("\u200b", "\u200c", "\u200d", "\u2060", "\ufeff"):
            with self.subTest(codepoint=hex(ord(char))):
                with self.assertRaises(ValueError):
                    validate_prompt_arguments({
                        "theme": "動画制作",
                        "video_path": "C:\\exports\\clip" + char + ".mp4",
                    })

    def test_normal_japanese_and_emoji_filename_remains_valid(self):
        args = {"theme": "動画制作", "video_path": "C:\\exports\\動画🎬.mp4"}
        self.assertEqual(validate_prompt_arguments(args), args)

    def test_windows_component_length_counts_utf16_units(self):
        for name in ("a" * 256, "🎬" * 128, "a" * 255 + "🎬"):
            with self.subTest(length=len(name)):
                with self.assertRaises(ValueError):
                    validate_prompt_arguments({
                        "theme": "動画制作",
                        "video_path": "C:\\exports\\" + name + ".mp4",
                    })

    def test_windows_component_near_boundary_is_accepted(self):
        # Use a 254-unit component: a portable near-limit positive case.
        args = {"theme": "動画制作",
                "video_path": "C:\\exports\\" + "a" * 250 + ".mp4"}
        self.assertEqual(validate_prompt_arguments(args), args)

    def test_decomposed_windows_output_paths_are_rejected(self):
        for path in ("C:\\exports\\cafe\u0301.mp4",
                     "\\\\server\\share\\re\u0301sume.mp4"):
            with self.subTest(path=path):
                with self.assertRaises(ValueError):
                    validate_prompt_arguments({"theme": "動画制作", "video_path": path})

    def test_precomposed_unicode_output_paths_remain_valid(self):
        for path in ("C:\\exports\\café.mp4",
                     "\\\\server\\share\\résumé.mp4"):
            with self.subTest(path=path):
                args = {"theme": "動画制作", "video_path": path}
                self.assertEqual(validate_prompt_arguments(args), args)

    def test_theme_can_still_contain_format_characters(self):
        args = {"theme": "A\u200dB"}
        self.assertEqual(validate_prompt_arguments(args), args)


if __name__ == "__main__":
    unittest.main()
