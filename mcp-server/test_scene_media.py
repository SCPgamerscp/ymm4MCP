"""Visual coverage is independent of audio on a scene timeline."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class SceneMediaTests(unittest.IsolatedAsyncioTestCase):
    async def test_audio_does_not_hide_visual_gaps(self):
        items = {"items": [{"item_id": "v", "type": "VideoItem", "frame": 0,
                            "length": 10, "layer": 1},
                           {"item_id": "a", "type": "AudioItem", "frame": 5,
                            "length": 25, "layer": 2}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=items)):
            result = await server.dispatch({"action": "get_info", "sub_action": "scenes",
                                            "scene_ranges": [{"id": "intro", "start_frame": 0,
                                                              "end_frame": 30}]})
        scene = result["scenes"][0]
        self.assertEqual(scene["visual_covered_frames"], 10)
        self.assertEqual(scene["audio_covered_frames"], 25)
        self.assertEqual(scene["visual_uncovered_ranges"], [{"start_frame": 10, "end_frame": 30}])
