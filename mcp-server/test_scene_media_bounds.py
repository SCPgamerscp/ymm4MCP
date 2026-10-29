"""Regression tests for scene-local coverage bounds and malformed timeline data."""
import copy
import unittest

from scene_media import with_coverage


def scene(items=None, start=10, end=20):
    return {"start_frame": start, "end_frame": end, "items": items or []}


def item(kind="VideoItem", start=10, end=20):
    return {"type": kind, "visible_start_frame": start, "visible_end_frame": end}


class SceneCoverageBoundsTests(unittest.TestCase):
    def test_clip_both_ends(self):
        s = scene([item(start=-100, end=100)])
        with_coverage({"scenes": [s]})
        self.assertEqual(s["visual_covered_frames"], 10)
        self.assertEqual(s["visual_coverage_ratio"], 1)
        self.assertEqual(s["visual_uncovered_ranges"], [])

    def test_clip_before_and_after_without_negative_coverage(self):
        s = scene([item(start=-50, end=12), item(start=18, end=100)])
        with_coverage({"scenes": [s]})
        self.assertEqual(s["visual_covered_frames"], 4)
        self.assertEqual(s["visual_uncovered_ranges"], [{"start_frame": 12, "end_frame": 18}])

    def test_outside_items_do_not_cover_scene(self):
        s = scene([item(start=-50, end=0), item(start=30, end=50)])
        with_coverage({"scenes": [s]})
        self.assertEqual(s["visual_covered_frames"], 0)
        self.assertEqual(s["visual_coverage_ratio"], 0)
        self.assertEqual(s["visual_uncovered_ranges"], [{"start_frame": 10, "end_frame": 20}])

    def test_overlaps_count_once_and_audio_is_independent(self):
        s = scene([item(start=5, end=16), item(start=14, end=25), item("VoiceItem", 12, 15)])
        with_coverage({"scenes": [s]})
        self.assertEqual(s["visual_covered_frames"], 10)
        self.assertEqual(s["audio_covered_frames"], 3)
        self.assertAlmostEqual(s["audio_coverage_ratio"], 0.3)

    def test_empty_scene_rejected_before_division(self):
        with self.assertRaisesRegex(ValueError, "positive duration"):
            with_coverage({"scenes": [scene(start=10, end=10)]})

    def test_boolean_and_fractional_frame_rejected(self):
        for invalid in (True, 11.5, "11"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    with_coverage({"scenes": [scene([item(start=invalid)])]})

    def test_reversed_item_rejected(self):
        with self.assertRaisesRegex(ValueError, "visible item range"):
            with_coverage({"scenes": [scene([item(start=19, end=18)])]})

    def test_invalid_later_scene_does_not_partially_mutate_earlier_scene(self):
        first = scene([item(start=11, end=15)])
        second = scene([item("AudioItem", 15, 14)])
        index = {"scenes": [first, second]}
        before = copy.deepcopy(index)
        with self.assertRaises(ValueError):
            with_coverage(index)
        self.assertEqual(index, before)


if __name__ == "__main__":
    unittest.main()
