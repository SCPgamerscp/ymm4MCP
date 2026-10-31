"""Regression tests for consecutive zero-progress ASGI request frames."""
import asyncio
import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location("http_body_limit", Path(__file__).with_name("http_body_limit.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


async def dispatch(chunks):
    statuses = []
    events = iter(chunks)

    async def receive():
        return next(events)

    async def send(message):
        if message["type"] == "http.response.start":
            statuses.append(message["status"])

    async def app(scope, receive, send):
        message = await receive()
        assert message["body"] == b"x" * sum(bool(c["body"]) for c in chunks)
        await send({"type": "http.response.start", "status": 200, "headers": []})
        await send({"type": "http.response.body", "body": b"ok"})

    scope = {"type": "http", "method": "POST", "http_version": "1.1", "headers": []}
    await module.RequestBodyLimit(app)(scope, receive, send)
    return statuses


class EmptyFrameProgressTests(unittest.TestCase):
    def test_progress_resets_consecutive_empty_frame_budget(self):
        frames = []
        for _ in range(2):
            frames.extend({"type": "http.request", "body": b"", "more_body": True}
                          for _ in range(module.MAX_EMPTY_REQUEST_EVENTS))
            frames.append({"type": "http.request", "body": b"x", "more_body": True})
        frames.append({"type": "http.request", "body": b"", "more_body": False})
        self.assertEqual(asyncio.run(dispatch(frames)), [200])

    def test_consecutive_empty_frames_still_rejected(self):
        frames = [{"type": "http.request", "body": b"", "more_body": True}
                  for _ in range(module.MAX_EMPTY_REQUEST_EVENTS + 1)]
        self.assertEqual(asyncio.run(dispatch(frames)), [413])


if __name__ == "__main__":
    unittest.main()
