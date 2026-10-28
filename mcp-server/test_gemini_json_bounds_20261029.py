"""Regression coverage for bounded Gemini response JSON parsing (Issue #61)."""
import unittest

import gemini_video


class GeminiJsonBoundsTests(unittest.TestCase):
    def parse(self, value):
        return gemini_video._extract_json(value)

    def assert_invalid(self, value):
        parsed = self.parse(value)
        self.assertIn("_parse_error", parsed)
        self.assertEqual(parsed["events"], [])
        self.assertEqual(parsed["summary"], "")

    def test_rejects_129_digit_integer_in_ignored_metadata(self):
        body = '{"summary":"ok","events":[],"extra":' + "9" * 129 + "}"
        self.assert_invalid(body)

    def test_rejects_129_digit_negative_integer_in_event(self):
        body = ('{"summary":"ok","events":[{"time_sec":-'
                + "9" * 129 + '}]}')
        self.assert_invalid(body)

    def test_accepts_128_digit_integer_in_unused_metadata(self):
        body = '{"summary":"ok","events":[],"extra":' + "9" * 128 + "}"
        parsed = self.parse(body)
        self.assertNotIn("_parse_error", parsed)
        self.assertEqual(parsed["summary"], "ok")

    def test_rejects_excessive_nested_arrays(self):
        deep = "[" * 129 + "0" + "]" * 129
        self.assert_invalid('{"summary":"ok","events":[],"extra":' + deep + "}")

    def test_accepts_nesting_at_limit(self):
        # The root object counts as one level, followed by 127 arrays.
        nested = "[" * 127 + "0" + "]" * 127
        parsed = self.parse('{"summary":"ok","events":[],"extra":' + nested + "}")
        self.assertNotIn("_parse_error", parsed)

    def test_brackets_and_escaped_quotes_in_strings_do_not_increase_depth(self):
        summary = 'brackets { [ ] } and escaped quote: \\"'
        import json
        body = json.dumps({"summary": summary, "events": []})
        parsed = self.parse(body)
        self.assertNotIn("_parse_error", parsed)
        self.assertEqual(parsed["summary"], summary)

    def test_rejects_deeply_nested_objects(self):
        deep = '{"nested":' * 129 + "0" + "}" * 129
        self.assert_invalid('{"summary":"ok","events":[],"extra":' + deep + "}")


if __name__ == "__main__":
    unittest.main()
