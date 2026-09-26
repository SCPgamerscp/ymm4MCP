"""Preview QA boundaries, decode failures, and position restoration."""
import base64
from io import BytesIO
import unittest
from unittest.mock import AsyncMock, patch

from PIL import Image

import server
import visual_qa


def png(color):
    output = BytesIO()
    Image.new("RGB", (80, 45), color).save(output, format="PNG")
    return base64.b64encode(output.getvalue()).decode()


class VisualQaTests(unittest.IsolatedAsyncioTestCase):
    async def test_black_static_and_restore_position(self):
        images = {0: png("black"), 30: png("black"), 60: png("black"),
                  90: png("white")}

        async def seek(_path, body, **_kwargs):
            return {"success": True, "image": images.get(body["frame"], png("white"))}

        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "currentFrame": 15, "totalFrames": 120})), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=seek)) as post:
            result = await server.dispatch({"action": "visual_qa", "end_frame": 90,
                                            "black_as_error": True})
        self.assertFalse(result["passed"])
        self.assertTrue(result["restored_position"])
        self.assertEqual([i["code"] for i in result["issues"]], ["BLACK_FRAME", "STATIC_PREVIEW"])
        self.assertEqual(post.await_args_list[-1].args[1], {"frame": 15})

    async def test_failed_capture_is_not_reported_as_pass(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"currentFrame": 2})), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[{"success": False, "error": "missing"},
                                                                  {"success": True}])) as post:
            result = await server.dispatch({"action": "visual_qa", "end_frame": 0})
        self.assertFalse(result["success"])
        self.assertFalse(result["passed"])
        self.assertEqual(post.await_count, 2)

    async def test_rejects_too_many_samples_before_seek(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "visual_qa", "end_frame": 1200})
        get.assert_not_awaited()

    def test_invalid_image(self):
        with self.assertRaises(ValueError):
            visual_qa.thumbnail(base64.b64encode(b"not png").decode())


if __name__ == "__main__":
    unittest.main()
