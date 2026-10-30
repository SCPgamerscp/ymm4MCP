"""Regression tests for binary-search scene membership and coverage validation."""
import unittest

import scenes
from scene_media import with_coverage


def ranges(count=2, *, gap=0, duration=10):
    stride = duration + gap
    return [{"id": f"scene-{i}", "start_frame": i * stride,
             "end_frame": i * stride + duration} for i in range(count)]


def item(identity, frame, length, kind="VideoItem"):
    return {"item_id": identity, "frame": frame, "length": length,
            "layer": 1, "type": kind}


class SceneMembershipSearchTests(unittest.TestCase):
    def test_exact_end_boundary_is_exclusive(self):
        result = scenes.index([item("left", 0, 10), item("right", 10, 10)], ranges())
        self.assertEqual([[x["item_id"] for x in scene["items"]]
                          for scene in result["scenes"]], [["left"], ["right"]])
        self.assertEqual(result["unassigned_item_ids"], [])
        with_coverage(result)

    def test_item_only_in_gap_is_unassigned(self):
        result = scenes.index([item("gap", 5, 5)], ranges(gap=5, duration=5))
        self.assertEqual(result["unassigned_item_ids"], ["gap"])
        self.assertEqual([scene["item_count"] for scene in result["scenes"]], [0, 0])
        self.assertEqual([scene["visual_covered_frames"] for scene in
                          with_coverage(result)["scenes"]], [0, 0])

    def test_item_intersecting_both_sides_of_gap(self):
        result = scenes.index([item("bridge", 4, 8)], ranges(gap=5, duration=5))
        self.assertEqual([[x["visible_start_frame"], x["visible_end_frame"]]
                          for scene in result["scenes"] for x in scene["items"]],
                         [[4, 5], [10, 12]])
        self.assertEqual(result["scenes"][0]["continues_into_next_item_ids"], ["bridge"])
        self.assertEqual(result["scenes"][1]["continues_from_previous_item_ids"], ["bridge"])
        self.assertEqual([scene["visual_covered_frames"] for scene in
                          with_coverage(result)["scenes"]], [1, 2])

    def test_item_after_last_scene_is_unassigned(self):
        result = scenes.index([item("late", 25, 5)], ranges())
        self.assertEqual(result["unassigned_item_ids"], ["late"])
        self.assertTrue(all(not scene["items"] for scene in result["scenes"]))
        with_coverage(result)

    def test_sparse_200_scenes_preserve_membership_and_coverage(self):
        result = scenes.index(
            [item("first", 0, 5), item("last", 1990, 5),
             item("gap", 1005, 5)], ranges(200, gap=5, duration=5))
        self.assertEqual(result["unassigned_item_ids"], ["gap"])
        self.assertEqual([i for i, scene in enumerate(result["scenes"])
                          if scene["items"]], [0, 199])
        self.assertEqual(sum(s["visual_covered_frames"] for s in
                             with_coverage(result)["scenes"]), 10)

    def test_long_item_crosses_all_200_scenes(self):
        result = scenes.index([item("long", 0, 2000)],
                              ranges(200, gap=5, duration=5))
        self.assertEqual(sum(s["item_count"] for s in result["scenes"]), 200)
        self.assertEqual(sum(s["visual_covered_frames"] for s in
                             with_coverage(result)["scenes"]), 1000)

    def test_gap_ending_exactly_at_next_scene_start(self):
        result = scenes.index([item("gap", 5, 5)], ranges(gap=5, duration=5))
        self.assertEqual(result["unassigned_item_ids"], ["gap"])
        with_coverage(result)


if __name__ == "__main__":
    unittest.main()
