"""Bound scene-index construction before expensive cross-scene expansion."""
import unittest
from unittest.mock import patch

import scenes
from scene_media import MAX_TOTAL_SCENE_ITEMS, with_coverage


RANGES = [{"id": "first", "start_frame": 0, "end_frame": 10},
          {"id": "second", "start_frame": 10, "end_frame": 20}]


def item(identity, frame=0, length=20, kind="VideoItem"):
    return {"item_id": identity, "frame": frame, "length": length,
            "layer": 1, "type": kind}


class SceneIndexBudgetTests(unittest.TestCase):
    def test_rejects_oversized_input_before_processing_items(self):
        with self.assertRaisesRegex(ValueError, "bounded array"):
            scenes.index([None] * (MAX_TOTAL_SCENE_ITEMS + 1), RANGES)

    def test_rejects_oversized_single_scene_membership(self):
        with patch.object(scenes, "MAX_SCENE_ITEMS", 2):
            with self.assertRaisesRegex(ValueError, "membership budget"):
                scenes.index([item("a"), item("b"), item("c")], RANGES)

    def test_rejects_oversized_aggregate_membership_from_crossings(self):
        with patch.object(scenes, "MAX_TOTAL_SCENE_ITEMS", 3):
            with self.assertRaisesRegex(ValueError, "membership budget"):
                scenes.index([item("a"), item("b")], RANGES)

    def test_accepts_exact_membership_budget(self):
        with patch.object(scenes, "MAX_SCENE_ITEMS", 2), patch.object(
                scenes, "MAX_TOTAL_SCENE_ITEMS", 4):
            index = scenes.index([item("a"), item("b")], RANGES)
        self.assertEqual([row["item_count"] for row in index["scenes"]], [2, 2])
        self.assertEqual(with_coverage(index)["scenes"][0]["visual_covered_frames"], 10)

    def test_rejects_control_characters_in_item_type_before_indexing(self):
        for kind in ("VideoItem\n", "VideoItem\x7f", "VideoItem\x85"):
            with self.subTest(kind=kind), self.assertRaisesRegex(
                    ValueError, "item.type"):
                scenes.index([item("v", kind=kind)], RANGES)

    def test_rejects_bidi_spoofed_item_type_before_indexing(self):
        for kind in ("VideoItem\u202e", "VideoItem\u2066", "VideoItem\u200f"):
            with self.subTest(kind=kind), self.assertRaisesRegex(
                    ValueError, "item.type"):
                scenes.index([item("v", kind=kind)], RANGES)

    def test_allows_known_namespaced_types_and_safe_unknown_types(self):
        for kind in ("YukkuriMovieMaker.Items.VideoItem", "CustomItem"):
            with self.subTest(kind=kind):
                index = scenes.index([item("v", kind=kind)], RANGES)
                self.assertEqual(len(index["scenes"][0]["items"]), 1)


if __name__ == "__main__":
    unittest.main()
