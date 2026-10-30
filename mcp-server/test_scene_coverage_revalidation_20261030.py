"""Regression tests for scene QA revalidation and cross-scene evidence."""
import copy
import unittest

import scenes
from scene_media import with_coverage


def snapshot():
    return scenes.index(
        [{"item_id": "video", "frame": 0, "length": 30,
          "layer": 1, "type": "VideoItem"},
         {"item_id": "audio", "frame": 10, "length": 10,
          "layer": 2, "type": "AudioItem"}],
        [{"id": "first", "start_frame": 0, "end_frame": 20},
         {"id": "second", "start_frame": 20, "end_frame": 40}])


def single(item):
    return {"scenes": [{"start_frame": 0, "end_frame": 20,
                        "items": [item]}]}


class SceneCoverageRevalidationTests(unittest.TestCase):
    def test_rejects_bidi_spoofed_host_item_type(self):
        for kind in ("VideoItem\u202e", "AudioItem\u2066", "VideoItem\u200f"):
            with self.subTest(kind=kind), self.assertRaisesRegex(
                    ValueError, "item type"):
                with_coverage(single({"type": kind,
                                      "visible_start_frame": 0,
                                      "visible_end_frame": 20}))

    def test_clips_legacy_source_span_before_scene(self):
        for kind, covered in (("UnknownItem", 0), ("VideoItem", 10)):
            with self.subTest(kind=kind):
                result = with_coverage(single({"type": kind,
                                               "visible_start_frame": -1,
                                               "visible_end_frame": 10}))
                self.assertEqual(result["scenes"][0]["visual_covered_frames"],
                                 covered)

    def test_clips_legacy_source_span_after_scene(self):
        for kind, covered in (("UnknownItem", 0), ("AudioItem", 10)):
            with self.subTest(kind=kind):
                result = with_coverage(single({"type": kind,
                                               "visible_start_frame": 10,
                                               "visible_end_frame": 21}))
                self.assertEqual(result["scenes"][0]["audio_covered_frames"],
                                 covered)

    def test_accepts_source_spans_clipped_to_scene(self):
        item = {"item_id": "v", "type": "VideoItem", "frame": 0,
                "end_frame": 30, "visible_start_frame": 0,
                "visible_end_frame": 20}
        result = with_coverage(single(item))
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_rejects_forged_visual_covered_frames_on_recheck(self):
        index = with_coverage({"scenes": [{"start_frame": 0, "end_frame": 1,
                                           "items": [{"type": "VideoItem",
                                                      "visible_start_frame": 0,
                                                      "visible_end_frame": 1}]}]})
        index["scenes"][0]["visual_covered_frames"] = True
        with self.assertRaisesRegex(ValueError, "visual_covered_frames"):
            with_coverage(index)

    def test_rejects_forged_audio_covered_frames_on_recheck(self):
        index = with_coverage(snapshot())
        index["scenes"][0]["audio_covered_frames"] = 11
        with self.assertRaisesRegex(ValueError, "audio_covered_frames"):
            with_coverage(index)

    def test_rejects_forged_coverage_ratio_even_when_equal_as_boolean(self):
        for invalid in (True, 1, float("nan"), float("inf"), -1.0):
            with self.subTest(invalid=invalid):
                index = with_coverage(snapshot())
                index["scenes"][0]["visual_coverage_ratio"] = invalid
                with self.assertRaisesRegex(ValueError, "visual_coverage_ratio"):
                    with_coverage(index)

    def test_rejects_forged_audio_ratio(self):
        index = with_coverage(snapshot())
        index["scenes"][0]["audio_coverage_ratio"] = 1.0
        with self.assertRaisesRegex(ValueError, "audio_coverage_ratio"):
            with_coverage(index)

    def test_rejects_forged_derived_gap_endpoints(self):
        index = with_coverage(snapshot())
        index["scenes"][0]["audio_uncovered_ranges"][0]["start_frame"] = False
        with self.assertRaisesRegex(ValueError, "audio_uncovered_ranges"):
            with_coverage(index)

    def test_rejects_boolean_in_scene_summary_gap(self):
        index = {"scenes": [{"start_frame": 0, "end_frame": 2,
                             "items": [{"type": "VideoItem",
                                        "visible_start_frame": 1,
                                        "visible_end_frame": 2}],
                             "uncovered_ranges": [{"start_frame": False,
                                                   "end_frame": True}]}]}
        with self.assertRaisesRegex(ValueError, "uncovered_ranges"):
            with_coverage(index)

    def test_rejects_boolean_in_layer_covered_frames(self):
        index = scenes.index(
            [{"item_id": "one", "frame": 0, "length": 1,
              "layer": 1, "type": "VideoItem"}],
            [{"id": "one", "start_frame": 0, "end_frame": 1}])
        index["scenes"][0]["layer_occupancy"][0]["covered_frames"] = True
        with self.assertRaisesRegex(ValueError, "layer_occupancy"):
            with_coverage(index)

    def test_rejects_boolean_in_layer_gap(self):
        index = snapshot()
        index["scenes"][0]["layer_occupancy"][1]["uncovered_ranges"][0]["start_frame"] = False
        with self.assertRaisesRegex(ValueError, "layer_occupancy"):
            with_coverage(index)

    def test_rejects_missing_cross_scene_membership_without_continuation_fields(self):
        first = {"item_id": "v", "type": "VideoItem", "frame": 0,
                 "end_frame": 30, "visible_start_frame": 0,
                 "visible_end_frame": 20}
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20, "items": [first]},
            {"start_frame": 20, "end_frame": 40, "items": []}]}
        with self.assertRaisesRegex(ValueError, "missing from intersecting scene"):
            with_coverage(index)

    def test_accepts_complete_cross_scene_membership_without_continuation_fields(self):
        first = {"item_id": "v", "type": "VideoItem", "frame": 0,
                 "end_frame": 30, "visible_start_frame": 0,
                 "visible_end_frame": 20}
        second = {**first, "visible_start_frame": 20,
                  "visible_end_frame": 30}
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20, "items": [first]},
            {"start_frame": 20, "end_frame": 40, "items": [second]}]}
        result = with_coverage(index)
        self.assertEqual(result["scenes"][1]["visual_covered_frames"], 10)

    def test_valid_recheck_is_idempotent(self):
        index = with_coverage(snapshot())
        before = copy.deepcopy(index)
        self.assertIs(with_coverage(index), index)
        self.assertEqual(index, before)

    def test_invalid_later_derived_evidence_does_not_mutate_earlier_scene(self):
        index = with_coverage(snapshot())
        index["scenes"][1]["visual_coverage_ratio"] = -1.0
        before = copy.deepcopy(index)
        with self.assertRaisesRegex(ValueError, "visual_coverage_ratio"):
            with_coverage(index)
        self.assertEqual(index, before)


if __name__ == "__main__":
    unittest.main()
