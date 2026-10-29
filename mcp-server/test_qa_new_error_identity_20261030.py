"""Regression: new anonymous QA failures must not be masked by repaired ones (#61)."""
import unittest

import editing


def item(text, frame=0, layer=1, kind="TextItem"):
    return {"frame": frame, "length": 10, "layer": layer,
            "type": kind, "text": text}


class QaNewErrorIdentityTests(unittest.TestCase):
    def test_new_missing_expectation_stops_even_when_error_count_is_constant(self):
        selectors = [{"text": "first"}, {"text": "second"}]
        before = editing.validate_timeline(
            [item("second")], expected=selectors, include_gaps=False)
        after = editing.validate_timeline(
            [item("first")], expected=selectors, include_gaps=False)
        self.assertEqual((before["score"], after["score"]), (80, 80))
        self.assertEqual(before["issues"][0]["code"], "EXPECTED_NOT_FOUND")
        self.assertEqual(after["issues"][0]["code"], "EXPECTED_NOT_FOUND")
        self.assertEqual(before["issues"][0]["expected_index"], 0)
        self.assertEqual(after["issues"][0]["expected_index"], 1)
        result = editing.evaluate_qa_gate(after, [before])
        self.assertEqual(result["reason_code"], "QA_REGRESSED")
        self.assertEqual(result["decision"], "stop")
        self.assertEqual(result["suggested_action"], "consider_checkpoint_rollback")

    def test_same_missing_expectation_after_reorder_still_stalls(self):
        selectors = [{"text": "missing"}]
        before = editing.validate_timeline(
            [item("a", layer=1), item("b", layer=2)],
            expected=selectors, include_gaps=False)
        after = editing.validate_timeline(
            [item("b", layer=2), item("a", layer=1)],
            expected=selectors, include_gaps=False)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")

    def test_new_missing_subtitle_with_anonymous_voice_is_regression(self):
        before = editing.validate_timeline(
            [item("first line", kind="VoiceItem")],
            subtitle_layers=[9], include_gaps=False)
        after = editing.validate_timeline(
            [item("second line", kind="VoiceItem")],
            subtitle_layers=[9], include_gaps=False)
        self.assertEqual((before["score"], after["score"]), (80, 80))
        self.assertEqual(before["issues"][0]["code"], "SUBTITLE_MISSING")
        self.assertEqual(after["issues"][0]["code"], "SUBTITLE_MISSING")
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_REGRESSED")

    def test_moving_same_anonymous_voice_does_not_count_as_new_error(self):
        before = editing.validate_timeline(
            [item("same line", frame=0, kind="VoiceItem")],
            subtitle_layers=[9], include_gaps=False)
        after = editing.validate_timeline(
            [item("same line", frame=50, kind="VoiceItem")],
            subtitle_layers=[9], include_gaps=False)
        result = editing.evaluate_qa_gate(after, [before])
        self.assertEqual(result["reason_code"], "QA_ISSUES_REMAIN")
        self.assertEqual(result["decision"], "repair")


if __name__ == "__main__":
    unittest.main()
