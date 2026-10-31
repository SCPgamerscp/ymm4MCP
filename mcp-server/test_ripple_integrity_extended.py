"""Regression: ripple verification must never bless collateral host mutations."""
import copy
import math
import unittest
from ripple_verification import verify_ripple_result


def item(identity, frame, **kwargs):
    return {"item_id": identity, "frame": frame, "length": 10, "layer": 1,
            "revision": "r1", "end_frame": frame + 10,
            "type": "VideoItem", "source_path": "C:/a.mp4", **kwargs}


def snap(*items):
    return {"success": True, "count": len(items), "items": list(items)}


class RippleIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.before = snap(item("cut", 10), item("moved", 20,
            effects=[{"type": "fade", "opacity": 0.5}],
            keyframes={"Zoom": [{"frame": 0, "value": 1.0}]},
            volume=0.25, source_offset=7), item("untouched", 0,
            effects=[{"type": "blur", "radius": 3}]))
        self.after = snap(item("moved", 10,
            effects=[{"type": "fade", "opacity": 0.5}],
            keyframes={"Zoom": [{"frame": 0, "value": 1.0}]},
            volume=0.25, source_offset=7), item("untouched", 0,
            effects=[{"type": "blur", "radius": 3}]))
        self.revisions = {"cut": "r1", "moved": "r1"}

    def check(self, after=None, before=None):
        return verify_ripple_result(self.before if before is None else before,
            self.after if after is None else after, start_frame=10, end_frame=20,
            layers=[1], expected_revisions=self.revisions)

    def test_valid_move_with_derived_end_passes(self):
        self.assertTrue(self.check()["passed"])

    def test_effect_mutation_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["effects"][0]["opacity"] = 0.75
        self.assertIn("effects", [i.get("field") for i in self.check(a)["issues"]])

    def test_keyframe_mutation_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["keyframes"]["Zoom"][0]["value"] = 1.5
        self.assertIn("keyframes", [i.get("field") for i in self.check(a)["issues"]])

    def test_audio_level_mutation_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["volume"] = 0.75
        self.assertIn("volume", [i.get("field") for i in self.check(a)["issues"]])

    def test_source_offset_mutation_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["source_offset"] = 8
        self.assertIn("source_offset", [i.get("field") for i in self.check(a)["issues"]])

    def test_unaffected_effect_mutation_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][1]["effects"][0]["radius"] = 4
        self.assertFalse(self.check(a)["passed"])

    def test_missing_field_detected(self):
        a = copy.deepcopy(self.after)
        del a["items"][0]["volume"]
        self.assertIn("volume", [i.get("field") for i in self.check(a)["issues"]])

    def test_new_field_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["muted"] = True
        self.assertIn("muted", [i.get("field") for i in self.check(a)["issues"]])

    def test_bool_numeric_spoof_detected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["source_offset"] = True
        self.assertFalse(self.check(a)["passed"])

    def test_stale_derived_end_rejected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["end_frame"] = 30
        self.assertEqual(self.check(a)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_missing_derived_end_detected(self):
        a = copy.deepcopy(self.after)
        del a["items"][0]["end_frame"]
        self.assertIn("end_frame", [i.get("field") for i in self.check(a)["issues"]])

    def test_nan_evidence_rejected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["volume"] = math.nan
        self.assertEqual(self.check(a)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_invalid_surrogate_rejected(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["source_path"] = "\ud800"
        self.assertEqual(self.check(a)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_revision_can_advance(self):
        a = copy.deepcopy(self.after)
        a["items"][0]["revision"] = "r2"
        self.assertTrue(self.check(a)["passed"])

    def test_inputs_unchanged(self):
        before, after = copy.deepcopy(self.before), copy.deepcopy(self.after)
        self.check()
        self.assertEqual(self.before, before)
        self.assertEqual(self.after, after)


if __name__ == "__main__":
    unittest.main()
