"""Read-only save/reopen acceptance test protocol, without running YMM4."""
import json
import os
from pathlib import Path
import tempfile
import unittest

import httpx

from check_keyframe_compat import ProbeError
from check_keyframe_persistence import capture, verify

BASE = "http://127.0.0.1:8765/api"
HEADERS = {"X-Ymm4-Token": "test-token"}


class FakeYmm4:
    def __init__(self, path):
        self.path = path
        self.saved = False
        self.persistent = True
        self.item_id = "native:test"
        self.frames = {"Zoom": [{"at": 0, "value": 1.0}, {"at": 24, "value": 1.25}],
                       "X": [{"at": 0, "value": 0.0}, {"at": 24, "value": 80.0}],
                       "Opacity": [{"at": 0, "value": 100.0}]}
        self.requests = []

    def __call__(self, request):
        self.requests.append(request)
        assert request.method == "GET", "persistence verification must never mutate YMM4"
        assert request.headers["X-Ymm4-Token"] == "test-token"
        if request.url.path == "/api/project":
            return httpx.Response(200, json={"success": True,
                "projectPath": str(self.path), "isSaved": self.saved})
        if request.url.path == "/api/items":
            return httpx.Response(200, json={"items": [{"item_id": self.item_id,
                "identity_persistent": self.persistent,
                "length": 100, "type": "ImageItem"}]})
        if request.url.path == "/api/items/keyframes":
            prop = request.url.params.get("prop")
            props = [prop] if prop else list(self.frames)
            return httpx.Response(200, json={"success": True,
                "inventory_complete": prop is None,
                "revision": "r1", "animations": [{"prop": p,
                    "keyframes": self.frames[p]} for p in props]})
        raise AssertionError(f"Unexpected path {request.url.path}")


class KeyframePersistenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.project = Path(self.tmp.name) / "throwaway.ymmp"
        self.project.write_bytes(b"original disposable project")
        self.snapshot = Path(self.tmp.name) / "capture.json"
        self.fake = FakeYmm4(self.project)
        self.client = httpx.Client(transport=httpx.MockTransport(self.fake), trust_env=False)
        self.addCleanup(self.client.close)

    def capture(self):
        capture(self.client, BASE, HEADERS, "native:test", 24,
                {"Zoom": 1.25, "X": 80.0}, self.snapshot)

    def save(self):
        self.project.write_bytes(b"updated disposable project")
        stat = self.project.stat()
        os.utime(self.project, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
        self.fake.saved = True

    def test_capture_then_save_and_reopen_verifies_full_inventory(self):
        self.capture()
        self.assertFalse(self.fake.saved)
        data = json.loads(self.snapshot.read_text())
        self.assertEqual(data["animations"]["Opacity"][0]["value"], 100.0)
        self.assertNotIn("test-token", self.snapshot.read_text())
        self.save()
        verify(self.client, BASE, HEADERS, self.snapshot)
        self.assertTrue(all(r.method == "GET" for r in self.fake.requests))

    def test_capture_refuses_preexisting_snapshot(self):
        self.capture()
        with self.assertRaisesRegex(ProbeError, "already exists"):
            self.capture()

    def test_capture_requires_unsaved_edit(self):
        self.fake.saved = True
        with self.assertRaisesRegex(ProbeError, "isSaved must be false"):
            self.capture()
        self.assertFalse(self.snapshot.exists())

    def test_capture_requires_persistent_item_identity(self):
        self.fake.persistent = False
        with self.assertRaisesRegex(ProbeError, "persistent ID"):
            self.capture()
        self.assertFalse(self.snapshot.exists())

    def test_capture_rejects_wrong_expected_value(self):
        self.fake.frames["X"][1]["value"] = 90.0
        with self.assertRaisesRegex(ProbeError, "expected 80"):
            self.capture()

    def test_verify_rejects_unsaved_or_not_written_project(self):
        self.capture()
        with self.assertRaisesRegex(ProbeError, "unsaved edits"):
            verify(self.client, BASE, HEADERS, self.snapshot)
        self.fake.saved = True
        with self.assertRaisesRegex(ProbeError, "not saved after capture"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_verify_detects_changed_other_animation(self):
        self.capture()
        self.save()
        self.fake.frames["Opacity"][0]["value"] = 80.0
        with self.assertRaisesRegex(ProbeError, "inventory differs"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_verify_detects_changed_project_path(self):
        self.capture()
        other = Path(self.tmp.name) / "different.ymmp"
        other.write_bytes(b"different")
        self.fake.path = other
        self.fake.saved = True
        with self.assertRaisesRegex(ProbeError, "Project path changed"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_verify_rejects_tampered_snapshot_schema(self):
        self.capture()
        self.save()
        data = json.loads(self.snapshot.read_text())
        data["values"]["Zoom"] = float("inf")
        self.snapshot.write_text(json.dumps(data))
        with self.assertRaisesRegex(ProbeError, "invalid schema"):
            verify(self.client, BASE, HEADERS, self.snapshot)


if __name__ == "__main__":
    unittest.main()
