"""Reject ambiguous stable IDs in scene snapshots without blocking cross-scene items."""
import unittest
import scenes


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 30},
          {"id": "body", "start_frame": 30, "end_frame": 60}]


class SceneIdentityTests(unittest.TestCase):
    def test_duplicate_ids_within_scene_are_rejected(self):
        items = [{"item_id": "same", "frame": 0, "length": 10},
                 {"item_id": "same", "frame": 15, "length": 10}]
        with self.assertRaisesRegex(ValueError, "duplicate item_id"):
            scenes.index(items, RANGES)

    def test_duplicate_ids_across_scenes_are_rejected(self):
        items = [{"item_id": "same", "frame": 0, "length": 10},
                 {"item_id": "same", "frame": 40, "length": 10}]
        with self.assertRaisesRegex(ValueError, "duplicate item_id"):
            scenes.index(items, RANGES)

    def test_duplicate_unassigned_ids_are_rejected(self):
        items = [{"item_id": "same", "frame": 60, "length": 10},
                 {"item_id": "same", "frame": 80, "length": 10}]
        with self.assertRaisesRegex(ValueError, "duplicate item_id"):
            scenes.index(items, RANGES)

    def test_whitespace_only_id_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "item_id is required"):
            scenes.index([{"item_id": " \t", "frame": 0, "length": 10}], RANGES)

    def test_single_item_crossing_scenes_is_not_a_duplicate(self):
        result = scenes.index([{"item_id": "same", "frame": 25, "length": 10}], RANGES)
        self.assertEqual([scene["item_count"] for scene in result["scenes"]], [1, 1])
        self.assertEqual(result["scenes"][0]["continues_into_next_item_ids"], ["same"])
        self.assertEqual(result["scenes"][1]["continues_from_previous_item_ids"], ["same"])

    def test_distinct_ids_remain_valid(self):
        result = scenes.index([{"item_id": "a", "frame": 0, "length": 10},
                               {"item_id": "b", "frame": 30, "length": 10}], RANGES)
        self.assertEqual([scene["item_count"] for scene in result["scenes"]], [1, 1])


if __name__ == "__main__":
    unittest.main()
