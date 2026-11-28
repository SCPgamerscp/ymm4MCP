"""Mocked protocol regression tests; the real YMM4 test is opt-in only."""
import json
import unittest

import httpx

from check_keyframe_compat import ProbeError, run_probe


BASE = "http://127.0.0.1:8765/api"
HEADERS = {"X-Ymm4-Token": "fake-test-token"}
TIMELINE_ID = "timeline:" + "a" * 32
ITEM_ID = "native:test/with spaces"


class FakeYmm4:
    def __init__(self):
        self.frames = {"Zoom": [{"at": 0, "value": 1.0}],
                       "X": [{"at": 0, "value": 0.0}],
                       "Opacity": [{"at": 0, "value": 100.0}]}
        self.posts = []
        self.revision = 1
        self.fail_set = None
        self.fail_reedit = None
        self.bad_reedit_value = None
        self.unversioned_reedit = None
        self.timeout_after_failed_set = False
        self.omit_set_revision = False
        self.external_change_after_unversioned_set = False
        self.external_replace_after_unversioned_set = False
        self.external_change_after_failed_set = False
        self.external_replace_after_failed_set = False
        self.bad_value = None
        self.fail_remove = None
        self.corrupt_original_on_remove = None
        self.corrupt_other_on_remove = None
        self.corrupt_unrelated_on_remove = None
        self.item_length = 100
        self.item_type = "ImageItem"
        self.item_count = 1
        self.timeline_id = TIMELINE_ID
        self.inventory_complete = True
        self.force_revision_conflict = False
        self.external_change_before_cleanup = None
        self.external_change_during_preflight = False
        self.external_change_before_write = False
        self.external_change_before_second_write = False
        self.property_get_count = {}

    def __call__(self, request):
        assert request.headers.get("X-Ymm4-Token") == "fake-test-token"
        assert request.url.host == "127.0.0.1"
        if request.method == "GET":
            if request.url.path == "/api/items":
                return httpx.Response(200, json={"timeline_id": self.timeline_id, "items": [
                    {"item_id": ITEM_ID,
                     "length": self.item_length, "type": self.item_type,
                     "revision": f"r{self.revision}"}
                    for _ in range(self.item_count)]})
            prop = request.url.params.get("prop")
            if prop is None:
                assert request.url.params.get("strict_inventory") == "1"
                return httpx.Response(200, json={"success": True,
                                                 "item_id": ITEM_ID,
                                                 "inventory_complete": self.inventory_complete,
                                                 "supports_expected_timeline_id": True,
                                                 "animations": [
                    {"prop": name, "keyframes": [dict(f) for f in frames]}
                    for name, frames in self.frames.items()],
                    "revision": f"r{self.revision}"})
            self.property_get_count[prop] = self.property_get_count.get(prop, 0) + 1
            # Simulate a concurrent edit to an unrelated animation between
            # inventory snapshots and property-specific revision reads.
            if ((prop == "Zoom" and self.property_get_count[prop] == 1
                 and self.external_change_during_preflight)
                or (prop == "Zoom" and self.property_get_count[prop] == 2
                    and self.external_change_before_write)
                or (prop == "X" and self.property_get_count[prop] == 2
                    and self.external_change_before_second_write)):
                self.frames["Opacity"][0]["value"] += 1
                self.revision += 1
            if (prop == self.external_change_before_cleanup
                    and self.property_get_count[prop] == 5):
                # Simulate a second editor replacing the just-inserted frame
                # between probe readback and the cleanup GET.
                self.frames[prop][-1]["value"] += 50
                self.revision += 1
            if prop not in self.frames:
                return httpx.Response(200, json={"success": True, "item_id": ITEM_ID, "animations": [],
                                                  "revision": f"r{self.revision}"})
            return httpx.Response(200, json={"success": True, "item_id": ITEM_ID, "animations": [
                {"prop": prop, "keyframes": [dict(f) for f in self.frames[prop]]}],
                "revision": f"r{self.revision}"})
        body = json.loads(request.content)
        self.posts.append(body)
        if body.get("expected_timeline_id", self.timeline_id) != self.timeline_id:
            return httpx.Response(200, json={"success": False,
                                              "error_code": "KEYFRAME_TIMELINE_CHANGED"})
        prop, at, action = body["prop"], body["at"], body["action"]
        if body["expected_revision"] != f"r{self.revision}":
            return httpx.Response(200, json={"success": False,
                                              "error_code": "REVISION_CONFLICT"})
        if action == "set":
            if self.force_revision_conflict:
                return httpx.Response(200, json={"success": False,
                                                  "error_code": "REVISION_CONFLICT"})
            value = body["value"]
            existing = [f for f in self.frames[prop] if f["at"] == at]
            is_reedit = bool(existing)
            if prop == self.bad_value or (is_reedit and prop == self.bad_reedit_value):
                value += 10
            if existing:
                existing[0]["value"] = value
            else:
                self.frames[prop].append({"at": at, "value": value})
            self.revision += 1
            if is_reedit and prop == self.fail_reedit:
                return httpx.Response(200, json={"success": False,
                                                  "error_code": "KEYFRAME_METHOD_UNAVAILABLE"})
            if is_reedit and prop == self.unversioned_reedit:
                return httpx.Response(200, json={"success": True})
            if prop == self.fail_set:
                # The host can mutate and then report failure. Another editor
                # may also modify this item before the probe can read it back.
                if self.external_change_after_failed_set:
                    self.frames["Opacity"][0]["value"] += 1
                    self.revision += 1
                if self.external_replace_after_failed_set:
                    self.frames[prop][-1]["value"] += 50
                    self.revision += 1
                if self.timeout_after_failed_set:
                    raise httpx.ReadTimeout("simulated lost response after mutation")
                return httpx.Response(200, json={"success": False,
                                                  "error_code": "KEYFRAME_METHOD_UNAVAILABLE"})
            if self.omit_set_revision:
                # A legacy host may mutate successfully but omit the revision.
                # Simulate an independent editor changing the item before the
                # probe reads it back, so cleanup must not adopt that revision.
                if self.external_change_after_unversioned_set:
                    self.frames["Opacity"][0]["value"] += 1
                    self.revision += 1
                if self.external_replace_after_unversioned_set:
                    self.frames[prop][-1]["value"] += 50
                    self.revision += 1
                return httpx.Response(200, json={"success": True})
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
        return httpx.Response(200, json={"success": True,
                                          "revision": f"r{self.revision}"})


def probe(fake, at=24):
    with httpx.Client(transport=httpx.MockTransport(fake), trust_env=False) as client:
        run_probe(client, BASE, HEADERS, ITEM_ID, at,
                  {"Zoom": 1.25, "X": 80.0})


class LiveProbeProtocolTests(unittest.TestCase):
    def test_missing_timeline_identity_prevents_writes(self):
        fake = FakeYmm4()

        def old_host(request):
            response = fake(request)
            if request.url.path == "/api/items":
                payload = response.json()
                payload.pop("timeline_id")
                return httpx.Response(200, json=payload)
            return response

        with self.assertRaisesRegex(ProbeError, "requires a live timeline identity"):
            probe(old_host)
        self.assertEqual(fake.posts, [])

    def test_timeline_identity_without_host_write_guard_prevents_writes(self):
        fake = FakeYmm4()

        def host_without_write_guard(request):
            response = fake(request)
            if request.url.params.get("strict_inventory") == "1":
                payload = response.json()
                payload.pop("supports_expected_timeline_id")
                return httpx.Response(200, json=payload)
            return response

        with self.assertRaisesRegex(ProbeError, "requires host support for timeline-bound keyframe writes"):
            probe(host_without_write_guard)
        self.assertEqual(fake.posts, [])

    def test_timeline_switch_at_insert_reedit_or_cleanup_is_rejected_by_host(self):
        for switch_at in (0, 1, 2):
            with self.subTest(switch_at=switch_at):
                fake = FakeYmm4()
                active_before = None

                def switched_at_post(request):
                    nonlocal active_before
                    if request.method == "POST" and len(fake.posts) == switch_at:
                        # A different timeline has the same item ID, animation
                        # state and revision when the queued write arrives.
                        fake.timeline_id = "timeline:" + "b" * 32
                        active_before = json.dumps(fake.frames)
                    return fake(request)

                with self.assertRaisesRegex(ProbeError, "KEYFRAME_TIMELINE_CHANGED"):
                    probe(switched_at_post)
                self.assertEqual(len(fake.posts), switch_at + 1)
                self.assertEqual([post["expected_timeline_id"] for post in fake.posts],
                                 [TIMELINE_ID] * (switch_at + 1))
                self.assertEqual(json.dumps(fake.frames), active_before)

    def test_incomplete_item_inventory_stops_before_keyframe_writes(self):
        for invalid in (
            {"count": 2},
            {"count": True},
            {"success": False},
            {"error": "timeline unavailable"},
            {"error_code": "TIMELINE_UNAVAILABLE"},
        ):
            with self.subTest(invalid=invalid):
                fake = FakeYmm4()

                def incomplete_inventory(request):
                    response = fake(request)
                    if request.url.path == "/api/items":
                        return httpx.Response(200, json={**response.json(), **invalid})
                    return response

                with self.assertRaisesRegex(ProbeError, "Malformed timeline item inventory"):
                    probe(incomplete_inventory, at=None)
                self.assertEqual(fake.posts, [])

    def test_missing_or_mismatched_readback_id_stops_before_keyframe_writes(self):
        for property_read, returned_id in ((False, None), (True, None),
                                           (False, "native:other"), (True, "native:other")):
            with self.subTest(property_read=property_read, returned_id=returned_id):
                fake = FakeYmm4()

                def mismatched_readback(request):
                    response = fake(request)
                    if (request.method == "GET" and request.url.path == "/api/items/keyframes"
                            and ("prop" in request.url.params) == property_read):
                        payload = response.json()
                        if returned_id is None:
                            payload.pop("item_id", None)
                        else:
                            payload["item_id"] = returned_id
                        return httpx.Response(200, json=payload)
                    return response

                with self.assertRaisesRegex(ProbeError, "Keyframe readback did not identify the requested item"):
                    probe(mismatched_readback, at=None)
                self.assertEqual(fake.posts, [])

    def test_live_timeline_switch_stops_writes_even_with_matching_revisions(self):
        for at, switch_after, new_identity in (
            (24, 0, "timeline:" + "b" * 32),
            (None, 0, None),
            (24, 1, "timeline:" + "b" * 32),
            (None, 2, "timeline:" + "b" * 32),
        ):
            with self.subTest(at=at, switch_after=switch_after, new_identity=new_identity):
                fake = FakeYmm4()
                initial = "timeline:" + "a" * 32
                item_reads = 0

                def switched(request):
                    nonlocal item_reads
                    response = fake(request)
                    if request.method == "GET" and request.url.path == "/api/items":
                        item_reads += 1
                        identity = (new_identity if item_reads > 1 and len(fake.posts) >= switch_after
                                    else initial)
                        payload = response.json()
                        if identity is None:
                            payload.pop("timeline_id")
                        else:
                            payload["timeline_id"] = identity
                        return httpx.Response(200, json=payload)
                    return response

                with self.assertRaisesRegex(ProbeError, "live timeline changed or became unavailable"):
                    probe(switched, at=at)
                self.assertEqual([post["action"] for post in fake.posts], ["set"] * switch_after)

    def test_stable_live_timeline_allows_insert_reedit_and_cleanup(self):
        fake = FakeYmm4()

        def same_timeline(request):
            response = fake(request)
            if request.method == "GET" and request.url.path == "/api/items":
                return httpx.Response(200, json={**response.json(), "count": 1,
                                                 "timeline_id": "timeline:" + "a" * 32})
            return response

        probe(same_timeline, at=None)
        self.assertEqual([post["action"] for post in fake.posts], ["set", "set", "remove"] * 2)
        self.assertEqual([post["expected_timeline_id"] for post in fake.posts], [TIMELINE_ID] * 6)

    def test_last_frame_uses_live_item_length_for_every_write(self):
        fake = FakeYmm4()
        fake.item_length = 73
        probe(fake, at=None)
        self.assertEqual([p["at"] for p in fake.posts], [72] * 6)
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertEqual([f["at"] for f in fake.frames["X"]], [0])

    def test_last_frame_refuses_existing_frame_and_one_frame_item(self):
        for length, existing in ((1, False), (73, True)):
            with self.subTest(length=length):
                fake = FakeYmm4()
                fake.item_length = length
                if existing:
                    fake.frames["X"].append({"at": length - 1, "value": 2.0})
                with self.assertRaises(ProbeError):
                    probe(fake, at=None)
                self.assertEqual(fake.posts, [])

    def test_last_frame_refuses_resize_after_preflight(self):
        fake = FakeYmm4()
        item_reads = 0

        def resized(request):
            nonlocal item_reads
            if request.url.path == "/api/items":
                item_reads += 1
                if item_reads == 2:
                    fake.item_length += 10
                    fake.revision += 1
            return fake(request)

        with self.assertRaisesRegex(ProbeError, "item length or revision changed"):
            probe(resized, at=None)
        self.assertEqual(fake.posts, [])

    def test_success_sets_reads_and_removes_both_properties(self):
        fake = FakeYmm4()
        probe(fake)
        self.assertEqual(len(fake.posts), 6)
        self.assertEqual([p["action"] for p in fake.posts],
                         ["set", "set", "remove", "set", "set", "remove"])
        self.assertEqual([p["prop"] for p in fake.posts],
                         ["Zoom", "Zoom", "Zoom", "X", "X", "X"])
        self.assertAlmostEqual(fake.posts[1]["value"], 1.125)
        self.assertAlmostEqual(fake.posts[4]["value"], 72.0)
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertEqual([f["at"] for f in fake.frames["X"]], [0])

    def test_wrong_reedit_value_fails_and_removes_only_own_frame(self):
        fake = FakeYmm4()
        fake.bad_reedit_value = "Zoom"
        with self.assertRaisesRegex(ProbeError, "wrong value after re-edit"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts],
                         ["set", "set", "remove"])
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])

    def test_reedit_failure_after_mutation_refuses_stale_cleanup(self):
        fake = FakeYmm4()
        fake.fail_reedit = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*changed after readback"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "set"])
        self.assertEqual(len(fake.frames["Zoom"]), 2)

    def test_reedit_without_revision_refuses_stale_cleanup(self):
        fake = FakeYmm4()
        fake.unversioned_reedit = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*changed after readback"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "set"])
        self.assertEqual(len(fake.frames["Zoom"]), 2)

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

    def test_incomplete_animation_inventory_rejected_before_write(self):
        fake = FakeYmm4()
        fake.inventory_complete = False
        with self.assertRaisesRegex(ProbeError, "Incomplete animation inventory"):
            probe(fake)
        self.assertEqual(fake.posts, [])

    def test_unrelated_edit_during_initial_preflight_aborts_before_any_write(self):
        fake = FakeYmm4()
        fake.external_change_during_preflight = True
        with self.assertRaisesRegex(ProbeError, "changed during preflight"):
            probe(fake)
        self.assertEqual(fake.posts, [])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)

    def test_unrelated_edit_after_inventory_read_aborts_before_write(self):
        fake = FakeYmm4()
        fake.external_change_before_write = True
        with self.assertRaisesRegex(ProbeError, "changed before write"):
            probe(fake)
        self.assertEqual(fake.posts, [])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)

    def test_second_property_rechecks_full_inventory_before_write(self):
        fake = FakeYmm4()
        fake.external_change_before_second_write = True
        with self.assertRaisesRegex(ProbeError, "changed before write"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "set", "remove"])
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)

    def test_revision_conflict_does_not_remove_another_editors_frame(self):
        fake = FakeYmm4()
        fake.force_revision_conflict = True
        with self.assertRaisesRegex(ProbeError, "REVISION_CONFLICT"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set"])
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])

    def test_concurrent_change_after_readback_refuses_cleanup(self):
        fake = FakeYmm4()
        fake.external_change_before_cleanup = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*changed after readback"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "set"])
        self.assertEqual(fake.frames["Zoom"][-1], {"at": 24, "value": 51.125})

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

    def test_lost_post_response_cleans_up_only_identifiable_own_frame(self):
        fake = FakeYmm4()
        fake.fail_set = "Zoom"
        fake.timeout_after_failed_set = True
        with self.assertRaisesRegex(ProbeError, "simulated lost response"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "remove"])
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])

    def test_failed_post_with_concurrent_unrelated_edit_refuses_cleanup(self):
        fake = FakeYmm4()
        fake.fail_set = "Zoom"
        fake.external_change_after_failed_set = True
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*unacknowledged edit is ambiguous"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set"])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)
        self.assertEqual(fake.frames["Zoom"][-1], {"at": 24, "value": 1.25})

    def test_failed_post_with_concurrent_replacement_refuses_cleanup(self):
        fake = FakeYmm4()
        fake.fail_set = "Zoom"
        fake.external_replace_after_failed_set = True
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*unacknowledged edit is ambiguous"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set"])
        self.assertEqual(fake.frames["Zoom"][-1], {"at": 24, "value": 51.25})

    def test_unversioned_success_is_not_certified_and_is_cleaned_if_isolated(self):
        fake = FakeYmm4()
        fake.omit_set_revision = True
        with self.assertRaisesRegex(ProbeError, "mutation succeeded without revision"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set", "remove"])
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])
        self.assertEqual(fake.frames["Opacity"], [{"at": 0, "value": 100.0}])

    def test_unversioned_success_does_not_delete_concurrent_unrelated_edit(self):
        fake = FakeYmm4()
        fake.omit_set_revision = True
        fake.external_change_after_unversioned_set = True
        with self.assertRaisesRegex(
                ProbeError, "CLEANUP FAILED.*unacknowledged edit is ambiguous"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set"])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)
        self.assertEqual(fake.frames["Zoom"][-1], {"at": 24, "value": 1.25})

    def test_unversioned_success_does_not_delete_replaced_keyframe(self):
        fake = FakeYmm4()
        fake.omit_set_revision = True
        fake.external_replace_after_unversioned_set = True
        with self.assertRaisesRegex(
                ProbeError, "CLEANUP FAILED.*unacknowledged edit is ambiguous"):
            probe(fake)
        self.assertEqual([p["action"] for p in fake.posts], ["set"])
        self.assertEqual(fake.frames["Zoom"][-1], {"at": 24, "value": 51.25})

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
        self.assertEqual(len(fake.posts), 3)

    def test_cleanup_rejects_collateral_damage_to_other_animation(self):
        fake = FakeYmm4()
        fake.corrupt_other_on_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*X"):
            probe(fake)
        self.assertEqual([f["at"] for f in fake.frames["Zoom"]], [0])
        self.assertNotEqual(fake.frames["X"][0]["value"], 0.0)
        self.assertEqual(len(fake.posts), 3)

    def test_cleanup_rejects_damage_to_unrelated_animation(self):
        fake = FakeYmm4()
        fake.corrupt_unrelated_on_remove = "Zoom"
        with self.assertRaisesRegex(ProbeError, "CLEANUP FAILED.*Opacity"):
            probe(fake)
        self.assertEqual(fake.frames["Zoom"], [{"at": 0, "value": 1.0}])
        self.assertEqual(fake.frames["X"], [{"at": 0, "value": 0.0}])
        self.assertEqual(fake.frames["Opacity"][0]["value"], 101.0)
        self.assertEqual(len(fake.posts), 3)


if __name__ == "__main__":
    unittest.main()
