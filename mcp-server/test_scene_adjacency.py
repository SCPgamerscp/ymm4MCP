"""Regression coverage for adjacent-scene continuation and stable layer order."""

import unittest

import scenes


class SceneAdjacencyTests(unittest.TestCase):
    def test_single_scene_has_no_neighbor_continuations(self):
        result = scenes.index(
            [{"item_id": "tail", "frame": 5, "length": 20, "layer": 1}],
            [{"id": "only", "start_frame": 0, "end_frame": 10}],
        )
        scene = result["scenes"][0]
        self.assertTrue(scene["items"][0]["crosses_boundary"])
        self.assertEqual(scene["continues_from_previous_item_ids"], [])
        self.assertEqual(scene["continues_into_next_item_ids"], [])

    def test_first_scene_has_no_previous_even_for_earlier_item(self):
        result = scenes.index(
            [{"item_id": "lead", "frame": 5, "length": 10, "layer": 1}],
            [{"id": "first", "start_frame": 10, "end_frame": 20}],
        )
        self.assertEqual(result["scenes"][0]["continues_from_previous_item_ids"], [])

    def test_item_ending_in_gap_is_not_continuation(self):
        ranges = [{"id": "a", "start_frame": 0, "end_frame": 10},
                  {"id": "b", "start_frame": 20, "end_frame": 30}]
        result = scenes.index(
            [{"item_id": "gap", "frame": 5, "length": 10, "layer": 1}],
            ranges,
        )
        self.assertEqual(result["scenes"][0]["continues_into_next_item_ids"], [])
        self.assertEqual(result["scenes"][1]["continues_from_previous_item_ids"], [])

    def test_item_spanning_gap_is_continuation(self):
        ranges = [{"id": "a", "start_frame": 0, "end_frame": 10},
                  {"id": "b", "start_frame": 20, "end_frame": 30}]
        result = scenes.index(
            [{"item_id": "bridge", "frame": 5, "length": 20, "layer": 1}],
            ranges,
        )
        self.assertEqual(result["scenes"][0]["continues_into_next_item_ids"], ["bridge"])
        self.assertEqual(result["scenes"][1]["continues_from_previous_item_ids"], ["bridge"])

    def test_layer_occupancy_is_sorted_even_for_shuffled_input(self):
        items = [
            {"item_id": "three", "frame": 0, "length": 5, "layer": 3},
            {"item_id": "missing", "frame": 0, "length": 5},
            {"item_id": "one", "frame": 0, "length": 5, "layer": 1},
            {"item_id": "two", "frame": 0, "length": 5, "layer": 2},
        ]
        ranges = [{"id": "a", "start_frame": 0, "end_frame": 10}]
        expected = [1, 2, 3, None]
        for ordered_items in (items, list(reversed(items))):
            with self.subTest(item_order=[item["item_id"] for item in ordered_items]):
                result = scenes.index(ordered_items, ranges)
                self.assertEqual(
                    [entry["layer"] for entry in result["scenes"][0]["layer_occupancy"]],
                    expected,
                )


if __name__ == "__main__":
    unittest.main()
