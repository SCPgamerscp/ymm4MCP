"""Regression test for the bundled Gemini video analysis model selection."""
import json
import unittest
from pathlib import Path

import gemini_video


class BundledGeminiModelTests(unittest.TestCase):
    def test_bundled_default_uses_supported_stable_video_model(self):
        config = json.loads((Path(gemini_video.__file__).parent / "config.json").read_text(
            encoding="utf-8"))
        self.assertEqual(config["default_model"], "gemini-3.8-flash")
        self.assertEqual(gemini_video.DEFAULT_MODEL, config["default_model"])


if __name__ == "__main__":
    unittest.main()
