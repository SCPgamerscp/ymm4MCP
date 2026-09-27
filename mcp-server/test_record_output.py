"""Recording output must be bounded, valid, and isolated between calls."""
import base64
from pathlib import Path
import re
import unittest
from unittest.mock import AsyncMock, patch

import server


class RecordOutputTests(unittest.IsolatedAsyncioTestCase):
    async def test_separate_recordings_do_not_overwrite(self):
        wav = b"RIFF" + (4).to_bytes(4, "little") + b"WAVE"
        response = {"success": True, "audio": base64.b64encode(wav).decode()}
        paths = []
        try:
            with patch.object(server, "ymm4_post", AsyncMock(return_value=response)):
                for _ in range(2):
                    result = await server.dispatch_preview({"action": "record"})
                    self.assertFalse(result.isError)
                    path = re.search(r"(/[^\s]+\.wav)", result.content[-1].text).group(1)
                    paths.append(Path(path))
            self.assertNotEqual(paths[0], paths[1])
            self.assertEqual([path.read_bytes() for path in paths], [wav, wav])
            self.assertIn("audio", response)
        finally:
            for path in paths:
                path.unlink(missing_ok=True)

    async def test_invalid_recording_does_not_create_file(self):
        with patch.object(server, "ymm4_post", AsyncMock(return_value={
                "success": True, "audio": "not-base64"})):
            result = await server.dispatch_preview({"action": "record"})
        self.assertTrue(result.isError)
