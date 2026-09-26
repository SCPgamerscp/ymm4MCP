"""Final MP4 acceptance dispatch validation."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class ExportQaDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_expected_output_criteria(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"passed": True})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "export_qa", "path": "C:/完成/clip.mp4",
                                   "expected_duration_seconds": 120, "duration_tolerance_seconds": 0.5,
                                   "expected_width": 1920, "expected_height": 1080, "require_audio": True})
            get.assert_awaited_once_with("/media/export-qa?path=C%3A%2F%E5%AE%8C%E6%88%90%2Fclip.mp4&expected_duration_seconds=120&duration_tolerance_seconds=0.5&expected_width=1920&expected_height=1080&require_audio=true")

    async def test_invalid_criteria_never_reach_ymm4(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            for option in ({"path": "clip.mp4"}, {"path": "C:/clip.wav"},
                           {"path": "C:/clip.mp4", "expected_duration_seconds": 0},
                           {"path": "C:/clip.mp4", "expected_width": True},
                           {"path": "C:/clip.mp4", "require_audio": "true"}):
                with self.subTest(option=option), self.assertRaises(ValueError):
                    await server.dispatch({"action": "get_info", "sub_action": "export_qa", **option})
            get.assert_not_awaited()
