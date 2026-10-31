"""Regression tests for optional Gemini event metadata schema (#61)."""
import json
import unittest

import gemini_video


class GeminiEventMetadataTests(unittest.TestCase):
    def parse_event(self, event):
        return gemini_video._extract_json(json.dumps({
            "summary": "ok", "events": [event],
        }))

    def test_rejects_non_textual_event_metadata(self):
        for field, value in (
            ("type", 7), ("label", ["wrong"]), ("description", {"bad": 1}),
            ("audio", False),
        ):
            with self.subTest(field=field):
                result = self.parse_event({"time_sec": 1, field: value})
                self.assertIn("_parse_error", result)
                self.assertEqual(result["events"], [])

    def test_rejects_invalid_importance(self):
        for value in (True, False, None, "5", 0, 6, 3.5, -1):
            with self.subTest(value=value):
                result = self.parse_event({"time_sec": 1, "importance": value})
                self.assertIn("_parse_error", result)

    def test_accepts_valid_optional_fields(self):
        for importance in range(1, 6):
            with self.subTest(importance=importance):
                result = self.parse_event({
                    "time_sec": 1.5, "end_sec": 2.0,
                    "type": "action", "label": "jump",
                    "description": "a character jumps",
                    "audio": None, "importance": importance,
                })
                self.assertNotIn("_parse_error", result)
                self.assertEqual(result["events"][0]["importance"], importance)

    def test_accepts_audio_text_and_omitted_metadata(self):
        self.assertNotIn("_parse_error", self.parse_event({
            "time_sec": 0, "audio": "music",
        }))
        self.assertNotIn("_parse_error", self.parse_event({"time_sec": 0}))


if __name__ == "__main__":
    unittest.main()
