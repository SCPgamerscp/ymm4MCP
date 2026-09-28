"""Scene grouping is read-only and must not assign ambiguous boundary items silently."""
import unittest
from unittest.mock import AsyncMock, patch

import scenes
import server


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 30},
          {"id": "body", "start_frame": 30, "end_frame": 60}]
ITEMS = [{"item_id": "a", "frame": 10, "length": 10, "layer": 1, "type": "TextItem"},
         {"item_id": "b", "frame": 25, "length": 10, "layer": 2, "type": "AudioItem"},
         {"item_id": "c", "frame": 60, "length": 10, "layer": 3, "type": "ImageItem"}]


class SceneIndexTests(unittest.TestCase):
    def test_overlap_and_boundary_items(self):
        result = scenes.index(ITEMS, RANGES)
        self.assertEqual([s["item_count"] for s in result["scenes"]], [2, 1])
        self.assertTrue(result["scenes"][0]["items"][1]["crosses_boundary"])
        self.assertTrue(result["scenes"][1]["items"][0]["crosses_boundary"])
        self.assertEqual(result["scenes"][0]["items"][1]["visible_start_frame"], 25)
        self.assertEqual(result["scenes"][0]["items"][1]["visible_end_frame"], 30)
        self.assertEqual(result["scenes"][1]["items"][0]["relative_start_frame"], 0)
        self.assertEqual(result["scenes"][1]["items"][0]["relative_end_frame"], 5)
        self.assertEqual(result["unassigned_item_ids"], ["c"])
        self.assertEqual([(s["covered_frames"], s["uncovered_frames"])
                          for s in result["scenes"]], [(15, 15), (5, 25)])

    def test_multiple_layers_cover_same_frames_only_once(self):
        items = [{"item_id": "x", "frame": 0, "length": 20},
                 {"item_id": "y", "frame": 10, "length": 20},
                 {"item_id": "z", "frame": 5, "length": 10}]
        result = scenes.index(items, RANGES + [{"id": "empty", "start_frame": 60,
                                                "end_frame": 70}])
        self.assertEqual([(s["covered_frames"], s["uncovered_frames"])
                          for s in result["scenes"]], [(30, 0), (0, 30), (0, 10)])

    def test_rejects_overlap_and_duplicate_names(self):
        for ranges in ([RANGES[0], {"id": "body", "start_frame": 29, "end_frame": 60}],
                       [RANGES[0], {"id": "intro", "start_frame": 30, "end_frame": 60}]):
            with self.subTest(ranges=ranges), self.assertRaises(ValueError):
                scenes.index([], ranges)


class SceneDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_fetches_one_snapshot(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": ITEMS})) as get:
            result = await server.dispatch({"action": "get_info", "sub_action": "scenes",
                                            "scene_ranges": RANGES})
        get.assert_awaited_once_with("/items")
        self.assertEqual(result["scenes"][1]["items"][0]["item_id"], "b")

    async def test_invalid_range_is_rejected_before_request(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "get_info", "sub_action": "scenes",
                                       "scene_ranges": [RANGES[1], RANGES[0]]})
            get.assert_not_awaited()
