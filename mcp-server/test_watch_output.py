"""Watch recordings are validated and saved without overwriting earlier captures."""
import base64
from pathlib import Path
import re
import unittest
from unittest.mock import AsyncMock, patch

import server


class WatchOutputTests(unittest.IsolatedAsyncioTestCase):
    async def test_separate_watch_recordings(self):
        wav = b"RIFF" + (4).to_bytes(4, "little") + b"WAVE"
        response = {"success": True, "audio": {"data": base64.b64encode(wav).decode()}, "frames": []}
        paths = []
        try:
            with patch.object(server, "ymm4_post", AsyncMock(return_value=response)):
                for _ in range(2):
                    result = await server.dispatch_preview({"action": "watch"})
                    self.assertFalse(result.isError)
                    paths.append(Path(re.search(r"(/[\w./-]+\.wav)", result.content[-1].text).group(1)))
            self.assertNotEqual(paths[0], paths[1])
            self.assertEqual([p.read_bytes() for p in paths], [wav, wav])
        finally:
            for path in paths:
                path.unlink(missing_ok=True)

    async def test_invalid_watch_audio_is_rejected(self):
        for data in ("not-base64", base64.b64encode(b"not wav").decode(), "a" * 42_000_001):
            with self.subTest(data=data[:12]), patch.object(server, "ymm4_post", AsyncMock(return_value={
                    "success": True, "audio": {"data": data}, "frames": []})):
                result = await server.dispatch_preview({"action": "watch"})
            self.assertTrue(result.isError)
