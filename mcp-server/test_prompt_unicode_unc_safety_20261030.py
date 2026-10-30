"""Untrusted production prompts must not spoof Windows file identities."""
import unittest

from prompt_validation import validate_prompt_arguments


class PromptUnicodeUncSafetyTests(unittest.TestCase):
    def assert_bad_path(self, path):
        with self.subTest(path=repr(path)), self.assertRaisesRegex(ValueError, 'video_path'):
            validate_prompt_arguments({'theme': '実況', 'video_path': path})

    def test_unc_server_and_share_are_validated(self):
        for path in (r'\\server\..\clip.mp4', r'\\ser?ver\share\clip.mp4',
                     r'\\server\sha?re\clip.mp4', r'\\server\CON\clip.mp4',
                     r'\\server\.\clip.mp4', r'\\server\share.\clip.mp4'):
            self.assert_bad_path(path)

    def test_repeated_or_empty_path_components_are_rejected(self):
        for path in ('C:/clips//clip.mp4', 'C:/clips/\\clip.mp4',
                     r'\\server\share\\clip.mp4', 'C:/clips/clip.mp4\\'):
            self.assert_bad_path(path)

    def test_windows_superscript_device_names(self):
        for path in ('C:/clips/COM¹.mp4', 'C:/clips/COM².txt',
                     'C:/clips/LPT³.mp4', r'\\server\share\LPT¹.mp4'):
            self.assert_bad_path(path)

    def test_path_bidi_and_c1_controls(self):
        for path in ('C:/clips/abc\u202e4pm.mp4', 'C:/clips/a\u2066b.mp4',
                     'C:/clips/a\u0085b.mp4', 'C:/clips/a\ud800.mp4',
                     r'\\server\sha\u202ere\clip.mp4'.replace('\\u202e', '\u202e')):
            self.assert_bad_path(path)

    def test_noncharacters_and_unpaired_surrogates_in_theme(self):
        for theme in ('abc\ud800', 'abc\uffff', 'abc\ufdd0'):
            with self.subTest(theme=repr(theme)), self.assertRaisesRegex(ValueError, 'theme'):
                validate_prompt_arguments({'theme': theme})

    def test_duration_accepts_only_plain_json_integers_or_canonical_ascii(self):
        class PretendInteger:
            def __int__(self):
                raise AssertionError('should never call __int__')
            def __str__(self):
                raise AssertionError('should never call __str__')
        for bad in (PretendInteger(), '１', '١', '+1', '0001', '1.0',
                    '9' * 10_000, 10 ** 1000, True, 1.0):
            with self.subTest(type=type(bad).__name__), self.assertRaisesRegex(ValueError, 'duration_seconds'):
                validate_prompt_arguments({'theme': '実況', 'duration_seconds': bad})
        for valid in (1, '1', 86400, '86400'):
            self.assertEqual(validate_prompt_arguments({'theme': '実況', 'duration_seconds': valid})['duration_seconds'], valid)

    def test_legitimate_unicode_and_unc_paths_remain_valid(self):
        for path in ('C:/動画/🎬clip.mp4', r'\\server\share\動画.mp4',
                     'C:/clips/COM10.mp4', 'C:/clips/conductor.mp4'):
            with self.subTest(path=path):
                args = {'theme': '日本語 🧪 / العربية', 'video_path': path}
                self.assertIs(validate_prompt_arguments(args), args)


if __name__ == '__main__':
    unittest.main()
