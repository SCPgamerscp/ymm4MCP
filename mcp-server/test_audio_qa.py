"""Audio QA dispatch validation without YMM4."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class AudioQaDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_audio_qa_dispatch(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "audio_qa",
                                   "path": "C:/素材/voice.wav", "min_silence_seconds": 1.5})
            get.assert_awaited_once_with("/media/audio-qa?path=C%3A%2F%E7%B4%A0%E6%9D%90%2Fvoice.wav&min_silence_seconds=1.5")

    async def test_invalid_options_do_not_call_ymm4(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            for path, seconds in (("relative.wav", 2), ("C:/voice.wav", 0),
                                  ("C:/voice.wav", True), ("C:/voice.wav", float("nan"))):
                with self.subTest(path=path, seconds=seconds), self.assertRaises(ValueError):
                    await server.dispatch({"action": "get_info", "sub_action": "audio_qa",
                                           "path": path, "min_silence_seconds": seconds})
            get.assert_not_awaited()
