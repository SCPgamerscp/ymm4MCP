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
BEFORE_TIMELINE = "timeline:" + "a" * 32
REOPENED_TIMELINE = "timeline:" + "b" * 32


class FakeYmm4:
    def __init__(self, path):
        self.path = path
        self.saved = False
        self.persistent = True
        self.item_id = "native:test"
        self.length = 100
        self.item_revision = "r1"
        self.animation_item_id = self.item_id
        self.timeline_metadata = {}
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
                "length": self.length, "revision": self.item_revision, "type": "ImageItem"}],
                **self.timeline_metadata})
        if request.url.path == "/api/items/keyframes":
            prop = request.url.params.get("prop")
            props = [prop] if prop else list(self.frames)
            identity = {} if self.animation_item_id is None else {"item_id": self.animation_item_id}
            return httpx.Response(200, json={"success": True, **identity,
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

    def capture(self, at=24):
        capture(self.client, BASE, HEADERS, "native:test", at,
                {"Zoom": 1.25, "X": 80.0}, self.snapshot)

    def save(self):
        self.project.write_bytes(b"updated disposable project")
        stat = self.project.stat()
        os.utime(self.project, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
        self.fake.saved = True

    def test_capture_then_save_and_reopen_verifies_full_inventory(self):
        self.fake.timeline_metadata = {"timeline_id": BEFORE_TIMELINE}
        self.capture()
        self.assertFalse(self.fake.saved)
        data = json.loads(self.snapshot.read_text())
        self.assertEqual(data["animations"]["Opacity"][0]["value"], 100.0)
        self.assertEqual(data["timeline_id"], BEFORE_TIMELINE)
        self.assertNotIn("test-token", self.snapshot.read_text())
        self.save()
        self.fake.timeline_metadata["timeline_id"] = REOPENED_TIMELINE
        verify(self.client, BASE, HEADERS, self.snapshot)
        self.assertTrue(all(r.method == "GET" for r in self.fake.requests))

    def test_verify_rejects_saving_without_reopening_live_timeline(self):
        self.fake.timeline_metadata = {"timeline_id": BEFORE_TIMELINE}
        self.capture()
        self.save()
        with self.assertRaisesRegex(ProbeError, "original live timeline is still active"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_verify_rejects_unidentified_or_different_animation_readback(self):
        self.capture()
        self.save()
        for returned_id in (None, "native:other"):
            with self.subTest(returned_id=returned_id):
                self.fake.animation_item_id = returned_id
                with self.assertRaisesRegex(ProbeError, "Keyframe readback did not identify the requested item"):
                    verify(self.client, BASE, HEADERS, self.snapshot)

    def test_verify_rejects_missing_live_timeline_after_capture(self):
        self.fake.timeline_metadata = {"timeline_id": BEFORE_TIMELINE}
        self.capture()
        self.save()
        self.fake.timeline_metadata.clear()
        with self.assertRaisesRegex(ProbeError, "now missing"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_capture_rejects_a_null_live_timeline(self):
        self.fake.timeline_metadata = {"timeline_id": None}
        with self.assertRaisesRegex(ProbeError, "Live timeline identity is unavailable"):
            self.capture()
        self.assertFalse(self.snapshot.exists())

    def test_verify_rejects_timeline_switch_during_readback_even_with_same_revision(self):
        self.fake.timeline_metadata = {"timeline_id": BEFORE_TIMELINE}
        self.capture()
        self.save()
        self.fake.timeline_metadata["timeline_id"] = REOPENED_TIMELINE

        def switch_timeline(request):
            response = self.fake(request)
            if request.url.params.get("prop") == "X":
                self.fake.timeline_metadata["timeline_id"] = "timeline:" + "c" * 32
            return response

        with httpx.Client(transport=httpx.MockTransport(switch_timeline), trust_env=False) as client:
            with self.assertRaisesRegex(ProbeError, "live timeline changed during snapshot"):
                verify(client, BASE, HEADERS, self.snapshot)

    def test_atomic_save_between_capture_and_verify_is_allowed(self):
        self.capture()
        replacement = self.project.with_suffix(".tmp")
        replacement.write_bytes(b"updated disposable project")
        stat = self.project.stat()
        os.utime(replacement, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))
        replacement.replace(self.project)
        self.fake.saved = True
        verify(self.client, BASE, HEADERS, self.snapshot)

    def client_replacing_project_during_readback(self):
        original = self.project.stat()
        replacement = self.project.with_suffix(".tmp")
        replacement.write_bytes(b"x" * original.st_size)
        os.utime(replacement, ns=(original.st_atime_ns, original.st_mtime_ns))

        def replace_project(request):
            response = self.fake(request)
            if request.url.params.get("prop") == "X":
                replacement.replace(self.project)
                self.assertEqual(self.project.stat().st_mtime_ns, original.st_mtime_ns)
            return response

        return httpx.Client(transport=httpx.MockTransport(replace_project), trust_env=False)

    def test_same_timestamp_project_replacement_during_capture_is_rejected(self):
        with self.client_replacing_project_during_readback() as client:
            with self.assertRaisesRegex(ProbeError, "Project changed"):
                capture(client, BASE, HEADERS, "native:test", 24,
                        {"Zoom": 1.25, "X": 80.0}, self.snapshot)
        self.assertFalse(self.snapshot.exists())

    def test_same_timestamp_project_replacement_during_verify_is_rejected(self):
        self.capture()
        self.save()
        with self.client_replacing_project_during_readback() as client:
            with self.assertRaisesRegex(ProbeError, "Project changed"):
                verify(client, BASE, HEADERS, self.snapshot)

    def test_capture_refuses_preexisting_snapshot(self):
        self.capture()
        with self.assertRaisesRegex(ProbeError, "already exists"):
            self.capture()

    def test_last_frame_is_resolved_and_verified_after_reopen(self):
        self.fake.length = 25
        self.capture(at=None)
        data = json.loads(self.snapshot.read_text())
        self.assertTrue(data["last_frame"])
        self.assertEqual(data["at"], 24)
        self.save()
        verify(self.client, BASE, HEADERS, self.snapshot)
        self.fake.length += 1
        for prop in ("Zoom", "X"):
            self.fake.frames[prop].append({**self.fake.frames[prop][-1], "at": 25})
        with self.assertRaisesRegex(ProbeError, "last frame differs"):
            verify(self.client, BASE, HEADERS, self.snapshot)

    def test_explicit_frame_snapshot_from_older_probe_still_verifies(self):
        self.capture()
        data = json.loads(self.snapshot.read_text())
        data.pop("last_frame", None)
        self.snapshot.write_text(json.dumps(data))
        self.save()
        verify(self.client, BASE, HEADERS, self.snapshot)

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

    def test_capture_rejects_image_revision_mismatch(self):
        self.fake.item_revision = "r2"
        with self.assertRaisesRegex(ProbeError, "Image revision changed"):
            self.capture()
        self.assertFalse(self.snapshot.exists())

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

    def test_verify_rechecks_project_after_reading_matching_animations(self):
        self.capture()
        self.save()
        other = Path(self.tmp.name) / "different.ymmp"
        other.write_bytes(b"different")

        def switch_project(request):
            response = self.fake(request)
            if request.url.params.get("prop") == "X":
                self.fake.path = other
            return response

        with httpx.Client(transport=httpx.MockTransport(switch_project), trust_env=False) as client:
            with self.assertRaisesRegex(ProbeError, "Project changed during"):
                verify(client, BASE, HEADERS, self.snapshot)

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
