"""Regression tests for exact VoiceItem/TextItem identification in structural QA."""
import unittest

from editing import validate_timeline


def item(item_id, kind, layer, text="hello"):
    return {"item_id": item_id, "type": kind, "frame": 0, "length": 30,
            "layer": layer, "text": text}


class SubtitleTypeIdentificationTests(unittest.TestCase):
    def test_unrelated_voice_suffix_is_not_a_voice(self):
        result = validate_timeline(
            [item("unrelated", "NotVoiceItem", 1)], subtitle_layers=[9])
        self.assertTrue(result["passed"])
        self.assertFalse(any(issue["code"] == "SUBTITLE_MISSING"
                             for issue in result["issues"]))

    def test_unrelated_text_suffix_cannot_satisfy_voice(self):
        result = validate_timeline(
            [item("voice", "VoiceItem", 1),
             item("unrelated", "NotTextItem", 9)], subtitle_layers=[9])
        self.assertFalse(result["passed"])
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["SUBTITLE_MISSING"])

    def test_namespace_and_nested_clr_types_are_supported(self):
        result = validate_timeline(
            [item("voice", "YukkuriMovieMaker.Project.Items.VoiceItem", 1),
             item("subtitle", "Custom.Namespace.Outer+TextItem", 9)],
            subtitle_layers=[9])
        self.assertTrue(result["passed"])

    def test_short_type_names_are_supported(self):
        result = validate_timeline(
            [item("voice", "VOICE", 1), item("subtitle", "TEXT", 9)],
            subtitle_layers=[9])
        self.assertTrue(result["passed"])


if __name__ == "__main__":
    unittest.main()
