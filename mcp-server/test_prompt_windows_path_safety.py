"""Production prompt video paths must be valid ordinary Windows file paths."""
import unittest

from prompt_validation import validate_prompt_arguments


class WindowsPromptPathSafetyTests(unittest.TestCase):
    def assert_invalid(self, path):
        with self.assertRaisesRegex(ValueError, "video_path"):
            validate_prompt_arguments({"theme": "test", "video_path": path})

    def test_rejects_forbidden_windows_filename_characters(self):
        for path in ("C:/clips/part?.mp4", "C:/clips/name*.mp4",
                     "C:/clips/a:b.mp4", "C:/clips/clip|copy.mp4",
                     "C:/clips/<name>.mp4", 'C:/clips/"quote".mp4',
                     r"\\server\share\clip?.mp4"):
            with self.subTest(path=path):
                self.assert_invalid(path)

    def test_rejects_windows_device_names_in_any_component(self):
        for path in ("C:/CON", "C:/clips/NUL.mp4", "C:/AUX/file.mp4",
                     "C:/clips/com1.wav", "C:/LPT9.txt",
                     r"\\server\share\PRN.mp4"):
            with self.subTest(path=path):
                self.assert_invalid(path)

    def test_rejects_ambiguous_windows_path_components(self):
        for path in ("C:/clips/../clip.mp4", "C:/clips/./clip.mp4",
                     "C:/clips/file.mp4.", "C:/clips/folder /clip.mp4",
                     "C:/clips/file.mp4 "):
            with self.subTest(path=path):
                self.assert_invalid(path)

    def test_allows_normal_windows_and_unc_files(self):
        for path in ("C:/clips/conductor.mp4", "D:/media/com10.wav",
                     "C:/素材/動画.mp4", r"\\server\share\video.mp4"):
            with self.subTest(path=path):
                args = {"theme": "test", "video_path": path}
                self.assertIs(validate_prompt_arguments(args), args)


if __name__ == "__main__":
    unittest.main()
