"""Gemini event frames and explicit file-analysis options must be safe."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class GeminiFrameConversionTests(unittest.TestCase):
    def test_valid_frame_conversion(self):
        self.assertEqual(server._sec_to_frame(2.5, 30, 100), 175)
        self.assertEqual(server._sec_to_frame(0, 30, 0), 0)
        self.assertEqual(server._sec_to_frame(0, 240, 2_147_483_647), 2_147_483_647)

    def test_invalid_times_cannot_produce_frame_numbers(self):
        for invalid in (True, "1", -1, float("nan"), float("inf"),
                        float("-inf"), 10 ** 500):
            with self.subTest(invalid=repr(invalid)):
                self.assertIsNone(server._sec_to_frame(invalid, 30, 0))

    def test_overflow_and_invalid_base_or_fps_return_none(self):
        for args in ((1e100, 240, 0), (1, 240, 2_147_483_647),
                     (1, True, 0), (1, 0, 0), (1, 241, 0),
                     (1, 30, True), (1, 30, -1), (1, 30, 2_147_483_648)):
            with self.subTest(args=args):
                self.assertIsNone(server._sec_to_frame(*args))


class GeminiFileArgumentTests(unittest.IsolatedAsyncioTestCase):
    async def test_invalid_explicit_fps_and_base_frame_fail_before_upload(self):
        with patch.object(server.gemini_video, "is_available", return_value=(True, "ok")),              patch.object(server.gemini_video, "analyze_video_file") as analyze:
            for args in ({"base_frame": "bad"}, {"base_frame": True},
                         {"base_frame": -1}, {"base_frame": 2_147_483_648},
                         {"fps": True}, {"fps": "30"}, {"fps": 0}, {"fps": 241}):
                with self.subTest(args=args), self.assertRaises(ValueError):
                    await server.analyze_video({"source": "file", "video_path": "C:/clip.mp4", **args})
            analyze.assert_not_called()

    async def test_invalid_host_fps_uses_default(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"fps": True})):
            self.assertEqual(await server._get_fps_from_ymm4(), 30)
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"fps": 241})):
            self.assertEqual(await server._get_fps_from_ymm4(), 30)


if __name__ == "__main__":
    unittest.main()
