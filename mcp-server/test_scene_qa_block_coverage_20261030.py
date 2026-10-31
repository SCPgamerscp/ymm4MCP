"""Regression coverage for disconnected selected-block QA gates (#61)."""
import unittest

import scene_qa


def scene(identity, start, end):
    return {"id": identity, "start_frame": start, "end_frame": end}


def item(identity, start, length, kind="VideoItem"):
    return {"item_id": identity, "frame": start, "length": length,
            "type": kind, "layer": 1}


SCOPES = [scene("a", 0, 10), scene("b", 10, 20),
          scene("c", 30, 40), scene("d", 40, 50)]


class ContiguousBlockCoverageTests(unittest.TestCase):
    def inspect(self, items, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": SCOPES, "require_audio": False,
            "max_visual_gap_frames": 1000, **options,
        })

    @staticmethod
    def codes(result):
        return [row["code"] for row in result["issues"]]

    def test_opt_in_only_does_not_change_existing_result(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 20)])
        self.assertTrue(result["passed"])
        self.assertNotIn("contiguous_blocks", result["overall_coverage"]["visual"])

    def test_disconnected_blocks_are_reported_separately(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 20)],
                              max_visual_block_uncovered_frames=0)
        self.assertTrue(result["passed"])
        blocks = result["overall_coverage"]["visual"]["contiguous_blocks"]
        self.assertEqual([row["scene_ids"] for row in blocks],
                         [["a", "b"], ["c", "d"]])
        self.assertEqual([row["scoped_frames"] for row in blocks], [20, 20])
        self.assertEqual([row["frame_range"] for row in blocks],
                         [[0, 20], [30, 50]])

    def test_block_leading_gap_detects_second_block(self):
        result = self.inspect([item("a", 0, 20), item("c", 32, 18)],
                              max_visual_block_leading_uncovered_frames=1)
        self.assertIn("SCENE_VISUAL_BLOCK_LEADING_UNCOVERED", self.codes(result))
        issue = result["issues"][0]
        self.assertEqual(issue["scene_ids"], ["c", "d"])
        self.assertEqual(issue["frame_range"], [30, 50])
        self.assertEqual(issue["leading_uncovered_frames"], 2)

    def test_block_leading_gap_exact_limit_passes(self):
        result = self.inspect([item("a", 0, 20), item("c", 32, 18)],
                              max_visual_block_leading_uncovered_frames=2)
        self.assertTrue(result["passed"])

    def test_block_trailing_gap_detects_second_block(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 18)],
                              max_visual_block_trailing_uncovered_frames=1)
        self.assertIn("SCENE_VISUAL_BLOCK_TRAILING_UNCOVERED", self.codes(result))
        self.assertEqual(result["issues"][0]["trailing_uncovered_frames"], 2)

    def test_block_trailing_gap_exact_limit_passes(self):
        self.assertTrue(self.inspect([item("a", 0, 20), item("c", 30, 18)],
                                     max_visual_block_trailing_uncovered_frames=2)["passed"])

    def test_block_uncovered_budget_is_not_diluted_by_first_block(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 15)],
                              max_visual_block_uncovered_frames=4,
                              max_visual_overall_uncovered_frames=5)
        self.assertIn("SCENE_VISUAL_BLOCK_UNCOVERED", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_OVERALL_UNCOVERED", self.codes(result))
        self.assertEqual(result["issues"][0]["uncovered_frames"], 5)

    def test_block_uncovered_budget_exact_limit_passes(self):
        self.assertTrue(self.inspect([item("a", 0, 20), item("c", 30, 15)],
                                     max_visual_block_uncovered_frames=5)["passed"])

    def test_block_ratio_catches_weak_block_when_global_ratio_passes(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 15)],
                              min_visual_overall_coverage_ratio=0.8,
                              min_visual_block_coverage_ratio=0.8)
        self.assertIn("SCENE_VISUAL_BLOCK_COVERAGE_LOW", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_OVERALL_COVERAGE_LOW", self.codes(result))
        self.assertEqual(result["issues"][0]["coverage_ratio"], 0.75)

    def test_block_ratio_exact_threshold_passes(self):
        self.assertTrue(self.inspect([item("a", 0, 20), item("c", 30, 15)],
                                     min_visual_block_coverage_ratio=0.75)["passed"])

    def test_block_gap_count_detects_fragmented_second_block(self):
        result = self.inspect([item("a", 0, 20), item("c", 32, 5),
                               item("d", 40, 10)],
                              max_visual_block_gap_count=1)
        self.assertIn("SCENE_VISUAL_BLOCK_GAP_COUNT", self.codes(result))
        self.assertEqual(result["issues"][0]["continuous_gap_count"], 2)

    def test_block_gap_count_exact_threshold_passes(self):
        self.assertTrue(self.inspect([item("a", 0, 20), item("c", 32, 5),
                                      item("d", 40, 10)],
                                     max_visual_block_gap_count=2)["passed"])

    def test_block_empty_scene_count_fails_only_second_block(self):
        result = self.inspect([item("a", 0, 20), item("d", 40, 10)],
                              max_visual_block_empty_scene_count=0)
        self.assertIn("SCENE_VISUAL_BLOCK_EMPTY_SCENE_COUNT", self.codes(result))
        self.assertEqual(result["issues"][0]["scene_ids"], ["c", "d"])
        self.assertEqual(result["issues"][0]["empty_scene_count"], 1)

    def test_block_empty_scene_count_exact_threshold_passes(self):
        self.assertTrue(self.inspect([item("a", 0, 20), item("d", 40, 10)],
                                     max_visual_block_empty_scene_count=1)["passed"])

    def test_audio_block_limits_are_independent_of_visual(self):
        result = scene_qa.inspect(
            [item("visual", 0, 20), item("audio", 0, 20, "VoiceItem"),
             item("visual2", 30, 20), item("audio2", 32, 18, "VoiceItem")],
            {"scene_ranges": SCOPES, "max_audio_gap_frames": 1000,
             "max_visual_gap_frames": 1000,
             "max_audio_block_leading_uncovered_frames": 1,
             "max_visual_block_leading_uncovered_frames": 0})
        self.assertIn("SCENE_AUDIO_BLOCK_LEADING_UNCOVERED", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_BLOCK_LEADING_UNCOVERED", self.codes(result))

    def test_unselected_time_is_not_a_block_gap(self):
        result = self.inspect([item("a", 0, 20), item("c", 30, 20)],
                              max_visual_block_leading_uncovered_frames=0,
                              max_visual_block_trailing_uncovered_frames=0,
                              max_visual_block_uncovered_frames=0,
                              min_visual_block_coverage_ratio=1,
                              max_visual_block_gap_count=0)
        self.assertTrue(result["passed"])

    def test_invalid_integer_limits_are_rejected(self):
        keys = ("max_visual_block_leading_uncovered_frames",
                "max_visual_block_trailing_uncovered_frames",
                "max_visual_block_uncovered_frames",
                "max_visual_block_gap_count",
                "max_visual_block_empty_scene_count")
        for key in keys:
            for invalid in (-1, True, 1.5, "1", None):
                with self.subTest(key=key, invalid=invalid):
                    with self.assertRaises(ValueError):
                        self.inspect([], **{key: invalid})

    def test_invalid_ratio_limits_are_rejected(self):
        for invalid in (-0.1, 1.01, True, "0.5", None, float("nan"),
                        float("inf")):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    self.inspect([], min_visual_block_coverage_ratio=invalid)

    def test_disabled_audio_rejects_block_options(self):
        keys = ("max_audio_block_leading_uncovered_frames",
                "max_audio_block_trailing_uncovered_frames",
                "max_audio_block_uncovered_frames",
                "max_audio_block_gap_count",
                "max_audio_block_empty_scene_count",
                "min_audio_block_coverage_ratio")
        for key in keys:
            with self.subTest(key=key):
                with self.assertRaisesRegex(ValueError, "disabled audio"):
                    self.inspect([], **{key: 0})


if __name__ == "__main__":
    unittest.main()
