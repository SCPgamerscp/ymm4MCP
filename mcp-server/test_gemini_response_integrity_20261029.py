"""Regression tests for fail-closed Gemini QA response parsing (#61)."""
import unittest

import gemini_video


class GeminiResponseIntegrityTests(unittest.TestCase):
    def parse(self, body):
        return gemini_video._extract_json(body)

    def assert_rejected(self, body):
        result = self.parse(body)
        self.assertIn("_parse_error", result)
        self.assertEqual(result["events"], [])

    def test_duplicate_top_level_and_nested_keys_are_rejected(self):
        self.assert_rejected('{"summary":"ok","summary":"changed","events":[]}')
        self.assert_rejected(
            '{"summary":"ok","events":[],"metadata":{"safe":true,"safe":false}}'
        )

    def test_nonstandard_json_constants_are_rejected_everywhere(self):
        for token in ("NaN", "Infinity", "-Infinity"):
            with self.subTest(token=token):
                self.assert_rejected(
                    '{"summary":"ok","events":[],"unused":' + token + '}'
                )

    def test_overflowed_and_excessively_long_floats_are_rejected(self):
        for token in ("1e9999", "-1e9999", "0." + "1" * 129):
            with self.subTest(token=token[:30]):
                self.assert_rejected(
                    '{"summary":"ok","events":[],"unused":' + token + '}'
                )

    def test_unpaired_unicode_surrogates_in_values_and_keys_are_rejected(self):
        self.assert_rejected(r'{"summary":"\ud800","events":[]}')
        self.assert_rejected(r'{"summary":"ok","events":[],"\udfff":"bad"}')
        self.assert_rejected(
            r'{"summary":"ok","events":[],"nested":["\ud800"]}'
        )

    def test_valid_emoji_surrogate_pair_is_preserved(self):
        result = self.parse(r'{"summary":"\ud83d\ude00","events":[]}')
        self.assertNotIn("_parse_error", result)
        self.assertEqual(result["summary"], "😀")

    def test_oversized_response_is_rejected_before_json_parsing(self):
        body = '{"summary":"' + "x" * gemini_video.MAX_GEMINI_RESPONSE_CHARS + '","events":[]}'
        self.assert_rejected(body)

    def test_valid_fenced_json_and_chronological_sort_still_work(self):
        fence = chr(96) * 3
        body = (
            fence + 'json\n{"summary":"ok","events":['
            '{"time_sec":2,"label":"later"},'
            '{"time_sec":1.5,"label":"earlier"}]}\n' + fence
        )
        result = self.parse(body)
        self.assertNotIn("_parse_error", result)
        self.assertEqual(
            [event["label"] for event in result["events"]],
            ["earlier", "later"],
        )


if __name__ == "__main__":
    unittest.main()
