"""Reject hallucinated or non-finite Gemini timestamps before frame conversion."""
import json
import unittest

import gemini_video


class GeminiTimestampValidationTests(unittest.TestCase):
    def parse(self, events):
        return gemini_video._extract_json(json.dumps({
            "summary": "test", "events": events,
        }, allow_nan=True))

    def test_rejects_missing_negative_boolean_and_nonfinite_start(self):
        for event in ({}, {"time_sec": -1}, {"time_sec": True},
                      {"time_sec": "1.0"}, {"time_sec": float("nan")},
                      {"time_sec": float("inf")}, {"time_sec": float("-inf")}):
            with self.subTest(event=event):
                self.assertIn("_parse_error", self.parse([event]))

    def test_rejects_end_before_start_and_nonfinite_end(self):
        for end in (-1, 1.5, True, "2", float("inf"), float("nan")):
            with self.subTest(end=end):
                self.assertIn("_parse_error", self.parse([{
                    "time_sec": 2.0, "end_sec": end,
                }]))

    def test_accepts_valid_instant_and_interval(self):
        events = [{"time_sec": 0}, {"time_sec": 2.5, "end_sec": None},
                  {"time_sec": 3.0, "end_sec": 4.0}]
        self.assertEqual(self.parse(events)["events"], events)

    def test_event_count_is_bounded(self):
        self.assertIn("_parse_error", self.parse([{"time_sec": 0}] * 5001))


if __name__ == "__main__":
    unittest.main()
