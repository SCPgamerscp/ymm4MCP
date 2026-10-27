"""Mocked protocol regression tests; the real YMM4 test is opt-in only."""
import json
import unittest

import httpx

from check_keyframe_compat import ProbeError, run_probe


BASE = "http://127.0.0.1:8765/api"
HEADERS = {"X-Ymm4-Token": "fake-test-token"}


class FakeYmm4:
    def __init__(self):
        self.frames = {"Zoom": [{"at": 0, "value": 1.0}],
                       "X": [{"at": 0, "value": 0.0}],
                       "Opacity": [{"at": 0, "value": 100.0}]}
        self.posts = []
        self.revision = 1
        self.fail_set = None
        self.bad_value = None
        self.fail_remove = None
        self.corrupt_original_on_remove = None
        self.corrupt_other_on_remove = None
        self.corrupt_unrelated_on_remove = None
        self.item_length = 100
        self.item_type = "ImageItem"
        self.item_count = 1

    def __call__(self, request):
        assert request.headers.get("X-Ymm4-Token") == "fake-test-token"
        assert request.url.host == "127.0.0.1"
        if request.method == "GET":
            if request.url.path == "/api/items":
                return httpx.Response(200, json={"items": [
                    {"item_id": "native:test/with spaces",
                     "length": self.item_length, "type": self.item_type}
                    for _ in range(self.item_count)]})
            prop = request.url.params.get("prop")
            if prop is None:
                return httpx.Response(200, json={"success": True, "animations": [
                    {"prop": name, "keyframes": [dict(f) for f in frames]}
                    for name, frames in self.frames.items()],
                    "revision": f"r{self.revision}"})
            if prop not in self.frames:
                return httpx.Response(200, json={"success": True, "animations": [],
                                                  "revision": f"r{self.revision}"})
            return httpx.Response(200, json={"success": True, "animations": [
                {"prop": prop, "keyframes": [dict(f) for f in self.frames[prop]]}],
                "revision": f"r{self.revision}"})
        body = json.loads(request.content)
        self.posts.append(body)
        prop, at, action = body["prop"], body["at"], body["action"]
        if body["expected_revision"] != f"r{self.revision}":
            return httpx.Response(200, json={"success": False,
                                              "error_code": "REVISION_CONFLICT"})
        if action == "set":
            value = body["value"]
            if prop == self.bad_value:
                value += 10
            self.frames[prop].append({"at": at, "value": value})
            self.revision += 1
            if prop == self.fail_set:
                return httpx.Response(200, json={"success": False,
                                                  "error_code": "KEYFRAME_METHOD_UNAVAILABLE"})
        elif action == "remove":
            if prop != self.fail_remove:
                self.frames[prop] = [f for f in self.frames[prop] if f["at"] != at]
            if prop == self.corrupt_original_on_remove:
                self.frames[prop][0]["value"] += 1
            if prop == self.corrupt_other_on_remove:
                other = "X" if prop == "Zoom" else "Zoom"
                self.frames[other][0]["value"] += 1
            if prop == self.corrupt_unrelated_on_remove:
                self.frames["Opacity"][0]["value"] += 1
            self.revision += 1
        return httpx.Response(200, json={"success": True})


def probe(fake):
    with httpx.Client(transport=httpx.MockTransport(fake), trust_env=False) as client:
        run_probe(client, BASE, HEADERS, "native:test/with spaces", 24,
                  {"Zoom": 1.25, "X": 80.0})


class LiveProbeProtocolTests(unittest.TestCase):
    def test_success_sets_reads_and_removes_both_properties(self):
        fake = FakeYmm4()
        probe(fake)
        self.assertEqual(len(fake.posts), 4)
        self.assertEqual([p["action"] for p in fake.posts],
                         ["set", "remove", "set", "remove"])
        self.assertEqual([p["prop"] for p in fake.posts],
                         ["Zoom", "Zoom", "X", "X"])
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertEqual([f["at"] for f in fake.frames["X"]], [0])

    def test_existing_keyframe_aborts_before_any_write(self):
        fake = FakeYmm4()
        fake.frames["X"].append({"at": 24, "value": 42.0})
        with self.assertRaisesRegex(ProbeError, "already exists"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_missing_animation_aborts_before_any_write(self):
        fake = FakeYmm4()
        del fake.frames["X"]
        with self.assertRaisesRegex(ProbeError, "unavailable"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_out_of_bounds_frame_rejected_before_write(self):
        fake = FakeYmm4()
        fake.item_length = 24
        with self.assertRaisesRegex(ProbeError, "outside the target item length"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_non_image_item_rejected_before_write(self):
        fake = FakeYmm4()
        fake.item_type = "AudioItem"
        with self.assertRaisesRegex(ProbeError, "not an image"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_missing_or_ambiguous_target_rejected_before_write(self):
        for count in (0, 2):
            with self.subTest(count=count):
                fake = FakeYmm4()
                fake.item_count = count
                with self.assertRaisesRegex(ProbeError, "exactly once"):
                    probe(fake)
                self.assertEqual(fake.posts, [])

    def test_malformed_item_length_rejected_before_write(self):
        fake = FakeYmm4()
        fake.item_length = "100"
        with self.assertRaisesRegex(ProbeError, "outside the target item length"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_post_can_fail_after_mutation_and_still_clean_up(self):
        fake = FakeYmm4()
        fake.fail_set = "Zoom"
        with self.assertRaisesRegex(ProbeError, "KEYFRAME_METHOD_UNAVAILABLE"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "remove"])
        self.assertEqual(len(fake.frames["Zoom"]), 1)

    def test_readback_mismatch_cleans_up_and_fails(self):
        fake = FakeYmm4()
        fake.bad_value = "X"
        with self.assertRaisesRegex(ProbeError, "wrong value"):
            probe(fake)
        self.assertEqual(len(fake.frames["Zoom"]), 1)
        self.assertEqual(len(fake.frames["X"]), 1)

    def test_cleanup_failure_is_reported_not_hidden(self):
        fake = FakeYmm4()
        fake.fail_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED"):
            probe(fake)
        self.assertEqual(len(fake.frames["Zoom"]), 2)

    def test_cleanup_rejects_corruption_of_original_keyframe(self):
        fake = FakeYmm4()
        fake.corrupt_original_on_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*Zoom"):
            probe(fake)
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertNotEqual(fake.frames["Zoom"][0]["value"], 1.0)
        self.assertEqual(len(fake.posts), 2)

    def test_cleanup_rejects_collateral_damage_to_other_animation(self):
        fake = FakeYmm4()
        fake.corrupt_other_on_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*X"):
            probe(fake)
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertNotEqual(fake.frames["X"][0]["value"], 0.0)
        self.assertEqual(len(fake.posts), 2)

    def test_cleanup_rejects_damage_to_unrelated_animation(self):
        fake = FakeYmm4()
        fake.corrupt_unrelated_on_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*Opacity"):
            probe(fake)
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])
        self.assertEqual(fake.frames["X"], [{"at": 0, "value": 0.0}])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)
        self.assertEqual(len(fake.posts), 2)


if __name__ == "__main__":
    unittest.main()
