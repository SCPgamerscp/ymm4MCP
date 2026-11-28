"""Declarative EditPlan validation, diff, idempotency, and MCP dispatch. No YMM4 required."""
import unittest
from unittest.mock import AsyncMock, patch

import httpx
import editplan
import server


def sample_plan(**overrides):
    plan = {
        "project": {"fps": 30},
        "gap": 0,
        "scenes": [
            {
                "id": "intro",
                "items": [
                    {"id": "bg-1", "type": "video", "source": "C:/Videos/play.mp4", "layer": 0, "frame": 0},
                    {"id": "line-1", "type": "dialogue", "character": "ゆっくり霊夢", "text": "こんにちは", "layer": 7},
                    {"id": "cap-1", "type": "subtitle", "text": "こんにちは", "layer": 9, "length": 90},
                ],
            }
        ],
    }
    plan.update(overrides)
    return {"plan": plan, "idempotency_key": "episode-1"}


def one_text_plan(text="hello"):
    return {"plan": {"scenes": [{"id": "scene", "items": [
        {"id": "caption", "type": "text", "text": text, "frame": 0, "layer": 1, "length": 10}
    ]}]}, "idempotency_key": "edit-1"}


def one_text_record(revision="r1"):
    return {"idempotency_key": "edit-1", "plan_hash": editplan.plan_hash(
        editplan.parse_plan(one_text_plan())), "items": [
            {"id": "caption", "item_id": "native:caption", "revision": revision}]}


def one_video_plan():
    return {"plan": {"scenes": [{"id": "scene", "items": [
        {"id": "clip", "type": "video", "source": "C:/Videos/play.mp4",
         "frame": 0, "layer": 1, "length": 10}
    ]}]}}


def one_voice_plan(**fields):
    return {"plan": {"scenes": [{"id": "scene", "items": [
        {"id": "line", "type": "dialogue", "character": "ゆっくり霊夢",
         "text": "こんにちは", "layer": 7, **fields}
    ]}]}}


def voice_item(**fields):
    return {"item_id": "native:existing-line", "revision": "r1", "type": "VoiceItem",
            "frame": 0, "length": 80, "layer": 7, "text": "こんにちは", **fields}


def video_item(item_id="native:clip", **fields):
    return {"item_id": item_id, "revision": "r1", "type": "VideoItem", "frame": 0,
            "length": 10, "layer": 1, "source_path": r"c:\VIDEOS\PLAY.MP4",
            "text": "Visible clip name", **fields}


class ParsePlanTests(unittest.TestCase):
    def test_explicit_frames_cannot_escape_their_scene(self):
        text_item = {"id": "b", "type": "text", "text": "b", "layer": 1, "length": 10}
        first = {"id": "a", "duration_policy": "explicit", "duration": 30,
                 "items": [{**text_item, "id": "a"}]}
        for second in ({"id": "b", "items": [{**text_item, "frame": 5}]},
                       {"id": "b", "duration_policy": "explicit", "duration": 10,
                        "items": [{**text_item, "frame": 35}]}):
            with self.subTest(second=second), self.assertRaises(ValueError):
                editplan.parse_plan({"plan": {"scenes": [first, second]}})
        with self.assertRaises(ValueError):
            editplan.parse_plan({"plan": {"start_frame": 2147483640, "scenes": [
                {"id": "a", "duration_policy": "explicit", "duration": 20,
                 "items": [{**text_item, "length": 1}]}]}})

    def test_snapshot_hash_ignores_item_order_but_detects_revision_changes(self):
        items = [{"item_id": "a", "revision": "r1", "frame": 0},
                 {"item_id": "b", "revision": "r1", "frame": 10}]
        self.assertEqual(editplan.snapshot_hash(items), editplan.snapshot_hash(list(reversed(items))))
        self.assertNotEqual(editplan.snapshot_hash(items),
                            editplan.snapshot_hash([{**items[0], "revision": "r2"}, items[1]]))

    def test_assigns_sequential_frames_and_hashes_stably(self):
        parsed = editplan.parse_plan(sample_plan())
        items = editplan.flatten_items(parsed)
        self.assertEqual(items[0]["frame"], 0)
        self.assertEqual(items[0]["kind"], "video")
        self.assertEqual(items[1]["frame_source"], "sequential")
        self.assertGreater(items[1]["length"], 1)
        self.assertEqual(items[2]["kind"], "text")
        self.assertEqual(editplan.plan_hash(parsed), editplan.plan_hash(editplan.parse_plan(sample_plan())))

    def test_rejects_invalid_plans_before_any_diff(self):
        invalid = [
            {},
            {"plan": {"scenes": []}},
            {"plan": {"scenes": [{"id": "a", "items": [{"id": "x", "type": "laser"}]}]}},
            {"plan": {"scenes": [{"id": "a", "items": [{"id": "x", "type": "video", "source": "rel.mp4"}]}]}},
            {"plan": {"scenes": [{"id": "a", "items": [
                {"id": "x", "type": "image", "source": "C:/a.png", "layer": 0}]}]}},
            {"plan": {"scenes": [{"id": "a", "items": [
                {"id": "x", "type": "dialogue", "character": "A", "text": "hi", "extra": 1}]}]}},
            {"plan": {"scenes": [{"id": "a", "items": [
                {"id": "dup", "type": "text", "text": "a", "length": 1},
                {"id": "dup", "type": "text", "text": "b", "length": 1}]}]}},
            {"plan": {"scenes": [{"id": "a", "duration_policy": "explicit", "items": [
                {"id": "x", "type": "text", "text": "a", "length": 1}]}]}},
            {"idempotency_key": " padded", "plan": {"scenes": [{"id": "a", "items": [
                {"id": "x", "type": "text", "text": "a", "length": 1}]}]}},
            {"atomic_scenes": "true", "plan": {"scenes": [{"id": "a", "items": [
                {"id": "x", "type": "text", "text": "a", "length": 1}]}]}},
        ]
        for args in invalid:
            with self.subTest(args=args), self.assertRaises(ValueError):
                editplan.parse_plan(args)

    def test_atomic_scenes_defaults_true_and_is_not_hashed(self):
        parsed = editplan.parse_plan(sample_plan())
        self.assertTrue(parsed["atomic_scenes"])
        off = editplan.parse_plan({**sample_plan(), "atomic_scenes": False})
        self.assertFalse(off["atomic_scenes"])
        self.assertEqual(editplan.plan_hash(parsed), editplan.plan_hash(off))
        self.assertTrue(editplan.checkpoint_id_ok("cp_" + "a" * 32))
        self.assertFalse(editplan.checkpoint_id_ok("cp_ab"))
        self.assertFalse(editplan.checkpoint_id_ok("../cp_x"))
        self.assertEqual(editplan.parse_reason({"reason": "before apply"}), "before apply")
        with self.assertRaises(ValueError):
            editplan.parse_reason({"reason": "x" * 201})

    def test_checkpoint_id_requires_exact_ascii_token(self):
        for token in ("cp_" + "a" * 8, "cp_" + "Z9" * 32):
            with self.subTest(token=token):
                self.assertTrue(editplan.checkpoint_id_ok(token))
        for token in (
            "cp_" + "a" * 7,
            "cp_" + "a" * 65,
            "cp_" + "a" * 8 + "\n",
            "cp_" + "a" * 8 + "\r\n",
            "cp_" + "a" * 8 + "\t",
            "cp_" + "a" * 8 + "/",
            "cp_" + "a" * 7 + "é",
            None,
            123,
        ):
            with self.subTest(token=token):
                self.assertFalse(editplan.checkpoint_id_ok(token))

    def test_group_ops_keeps_scene_order(self):
        ops = [
            {"id": "a", "scene_id": "intro", "op": "add"},
            {"id": "b", "scene_id": "intro", "op": "keep"},
            {"id": "c", "scene_id": "body", "op": "add"},
        ]
        grouped = editplan.group_ops_by_scene(ops)
        self.assertEqual([g["id"] for g in grouped], ["intro", "body"])
        self.assertEqual([op["id"] for op in grouped[0]["ops"]], ["a", "b"])

    def test_explicit_scene_duration_and_unique_scene_ids(self):
        parsed = editplan.parse_plan({"plan": {"scenes": [
            {"id": "a", "duration_policy": "explicit", "duration": 300, "items": [
                {"id": "t1", "type": "text", "text": "a", "length": 30, "layer": 1}]},
            {"id": "b", "items": [
                {"id": "t2", "type": "text", "text": "b", "length": 10, "layer": 1}]},
        ]}})
        self.assertEqual(parsed["scenes"][1]["start_frame"], 300)
        with self.assertRaises(ValueError):
            editplan.parse_plan({"plan": {"scenes": [
                {"id": "a", "items": [{"id": "t1", "type": "text", "text": "a", "length": 1}]},
                {"id": "a", "items": [{"id": "t2", "type": "text", "text": "b", "length": 1}]},
            ]}})


class DiffAndReplayTests(unittest.TestCase):
    def test_empty_timeline_plans_adds_and_content_match_keeps(self):
        parsed = editplan.parse_plan(sample_plan())
        preview = editplan.diff_plan(parsed, [], None, ["ゆっくり霊夢"])
        self.assertTrue(preview["dry_run"])
        self.assertEqual(preview["added"], 3)
        self.assertEqual(preview["kept"], 0)
        self.assertEqual([op["kind"] for op in preview["ops"]], ["video", "voice", "text"])
        self.assertEqual(preview["ops"][1]["payload"]["character"], "ゆっくり霊夢")
        self.assertNotIn("length", preview["ops"][1]["payload"])

        current = [
            {"item_id": "native:bg", "type": "VideoItem", "layer": 0, "frame": 0, "length": 900,
             "text": "C:/Videos/play.mp4", "revision": "r0"},
            {"item_id": "native:line", "type": "VoiceItem", "layer": 7, "frame": 0, "length": 80,
             "text": "こんにちは", "character": "ゆっくり霊夢", "revision": "r1"},
        ]
        matched = editplan.diff_plan(parsed, current, None, ["ゆっくり霊夢"])
        self.assertEqual(matched["kept"], 2)
        self.assertEqual(matched["added"], 1)
        self.assertEqual(matched["ops"][2]["id"], "cap-1")

    def test_existing_voice_without_character_is_an_explicit_conflict(self):
        plan = editplan.parse_plan(one_voice_plan(frame=0))
        for item in (voice_item(), voice_item(character=""), voice_item(character=None)):
            with self.subTest(item=item):
                preview = editplan.diff_plan(plan, [item])
                self.assertFalse(preview["passed"])
                self.assertEqual(preview["added"], 0)
                self.assertEqual(preview["warnings"][0]["reason"], "VOICE_CHARACTER_UNAVAILABLE")

    def test_explicitly_repeated_voice_is_allowed_at_a_different_frame(self):
        preview = editplan.diff_plan(editplan.parse_plan(one_voice_plan(frame=100)), [voice_item()])
        self.assertTrue(preview["passed"])
        self.assertEqual(preview["added"], 1)

    def test_known_voice_character_is_checked_in_the_snapshot_hash(self):
        item = voice_item(character="ゆっくり霊夢")
        self.assertNotEqual(
            editplan.snapshot_hash([item]),
            editplan.snapshot_hash([{**item, "character": "ゆっくり魔理沙"}]))

    def test_bindings_and_idempotent_replay(self):
        parsed = editplan.parse_plan(sample_plan())
        record = {
            "idempotency_key": "episode-1",
            "plan_hash": editplan.plan_hash(parsed),
            "items": [
                {"id": "bg-1", "item_id": "native:bg", "revision": "r0"},
                {"id": "line-1", "item_id": "native:line", "revision": "r1"},
                {"id": "cap-1", "item_id": "native:cap", "revision": "r2"},
            ],
        }
        current = [
            {"item_id": "native:bg", "type": "VideoItem", "layer": 0, "frame": 0, "length": 900, "text": "x", "revision": "r0"},
            {"item_id": "native:line", "type": "VoiceItem", "layer": 7, "frame": 0, "length": 80, "text": "こんにちは", "revision": "r1"},
            {"item_id": "native:cap", "type": "TextItem", "layer": 9, "frame": 0, "length": 90, "text": "こんにちは", "revision": "r2"},
            {"item_id": "native:extra", "type": "AudioItem", "layer": 1, "frame": 0, "length": 30, "text": "bgm.wav"},
        ]
        replay = editplan.replay_record(parsed, current, record)
        self.assertTrue(replay["replayed"])
        self.assertEqual(replay["added"], 0)
        preview = editplan.diff_plan(parsed, current, record, ["ゆっくり霊夢"])
        self.assertEqual(preview["kept"], 3)
        self.assertEqual(preview["added"], 0)
        self.assertTrue(any(w["code"] == "EXTRA_ITEM" for w in preview["warnings"]))

        missing = [item for item in current if item["item_id"] != "native:cap"]
        self.assertIsNone(editplan.replay_record(parsed, missing, record))
        reconciled = editplan.diff_plan(parsed, missing, record, ["ゆっくり霊夢"])
        self.assertEqual(reconciled["added"], 1)
        self.assertEqual(reconciled["ops"][2]["id"], "cap-1")

    def test_unknown_character_is_a_structured_warning(self):
        parsed = editplan.parse_plan(sample_plan())
        preview = editplan.diff_plan(parsed, [], None, ["ゆっくり魔理沙"])
        self.assertFalse(preview["passed"])
        self.assertTrue(any(w["code"] == "CHARACTER_UNKNOWN" for w in preview["warnings"]))

    def test_overlap_inside_the_plan_is_reported(self):
        parsed = editplan.parse_plan({"plan": {"scenes": [{"id": "s", "items": [
            {"id": "a", "type": "text", "text": "one", "layer": 1, "frame": 0, "length": 50},
            {"id": "b", "type": "text", "text": "two", "layer": 1, "frame": 10, "length": 50},
        ]}]}})
        preview = editplan.diff_plan(parsed, [])
        self.assertTrue(any(w["code"] == "OVERLAP" for w in preview["warnings"]))

    def test_edit_state_is_a_single_scene(self):
        state = editplan.to_edit_state([
            {"item_id": "native:a", "type": "VoiceItem", "layer": 7, "frame": 3, "length": 10, "text": "hi"}
        ])
        self.assertEqual(state["scenes"][0]["id"], "timeline")
        self.assertEqual(state["scenes"][0]["items"][0]["type"], "voice")

    def test_changed_binding_is_not_replayed(self):
        plan = editplan.parse_plan(one_text_plan())
        current = [{"item_id": "native:caption", "revision": "r2"}]
        self.assertIsNone(editplan.replay_record(plan, current, one_text_record()))
        conflict = editplan.binding_conflict(plan, current, one_text_record())
        self.assertEqual(conflict["error_code"], "EDIT_STATE_CONFLICT")
        self.assertEqual(conflict["conflicts"][0]["id"], "caption")

    def test_missing_binding_can_be_reconciled(self):
        plan = editplan.parse_plan(one_text_plan())
        self.assertIsNone(editplan.binding_conflict(plan, [], one_text_record()))
        self.assertEqual(editplan.diff_plan(plan, [], one_text_record())["added"], 1)

    def test_unbound_content_at_wrong_frame_is_a_conflict(self):
        plan = editplan.parse_plan(one_text_plan())
        current = [{"item_id": "native:caption", "revision": "r1", "type": "TextItem",
                    "text": "hello", "frame": 5, "layer": 1, "length": 10}]
        preview = editplan.diff_plan(plan, current)
        self.assertFalse(preview["passed"])
        self.assertEqual(preview["warnings"][0]["code"], "CONTENT_STATE_CONFLICT")

    def test_windows_source_path_is_reused_despite_slash_case_and_display_text(self):
        parsed = editplan.parse_plan(one_video_plan())
        for host_field in ("source_path", "path", "source", "text"):
            host = video_item()
            if host_field != "source_path":
                host[host_field] = host.pop("source_path")
            with self.subTest(host_field=host_field):
                preview = editplan.diff_plan(parsed, [host])
                self.assertTrue(preview["passed"])
                self.assertEqual(preview["added"], 0)
                self.assertEqual(preview["ops"][0]["item_id"], "native:clip")

    def test_posix_media_identity_remains_case_sensitive(self):
        item = video_item(source_path="/tmp/Media/clip.mp4")
        self.assertNotEqual(editplan.current_fingerprint(item),
                            editplan.current_fingerprint({**item, "source_path": "/tmp/media/clip.mp4"}))

    def test_repeated_media_is_reused_only_if_explicit_location_selects_one_item(self):
        parsed = editplan.parse_plan(one_video_plan())
        existing = video_item()
        for unrelated in (video_item("native:later", frame=12), video_item("native:longer", length=15)):
            with self.subTest(unrelated=unrelated["item_id"]):
                preview = editplan.diff_plan(parsed, [unrelated, existing])
                self.assertTrue(preview["passed"])
                self.assertEqual(preview["added"], 0)
                self.assertEqual(preview["ops"][0]["item_id"], existing["item_id"])

    def test_implicit_operation_cannot_hide_an_ambiguous_explicit_match(self):
        request = one_video_plan()
        planned = request["plan"]["scenes"][0]["items"]
        planned[0].pop("frame")
        planned.append({**planned[0], "id": "later-clip", "frame": 20})
        existing = [video_item("native:one", frame=20), video_item("native:two", frame=20)]

        preview = editplan.diff_plan(editplan.parse_plan(request), existing)
        conflicts = [warning for warning in preview["warnings"]
                     if warning["code"] == "CONTENT_STATE_CONFLICT"]
        self.assertFalse(preview["passed"])
        self.assertEqual([(warning["id"], warning["candidate_item_ids"]) for warning in conflicts],
                         [("later-clip", ["native:one", "native:two"])])

    def test_added_item_verification_checks_identity_and_revision(self):
        detail = {"id": "caption", "op": "add", "item_id": "native:caption",
                  "result": {"revision": "r1", "frame": 0, "layer": 1, "length": 10}}
        self.assertEqual(editplan.verify_applied_items([detail], []), [
            {"id": "caption", "item_id": "native:caption", "reason": "ITEM_NOT_UNIQUE"}])
        current = [{"item_id": "native:caption", "revision": "r2", "frame": 0,
                    "layer": 1, "length": 10}]
        self.assertEqual(editplan.verify_applied_items([detail], current)[0]["fields"], ["revision"])


class EditPlanDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_reconcile_dry_run_never_silently_applies(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaisesRegex(ValueError, "plan_edit"):
                await server.dispatch({"action": "reconcile_edit", **one_text_plan(),
                                       "dry_run": True})
        get.assert_not_awaited()
        post.assert_not_awaited()

    async def test_plan_edit_rejects_conflicting_dry_run_false(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaisesRegex(ValueError, "apply_edit"):
                await server.dispatch({"action": "plan_edit", **one_text_plan(),
                                       "dry_run": False})
            get.assert_not_awaited()

    async def test_expected_snapshot_hash_stops_stale_apply_before_post(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"characters": [{"name": "ゆっくり霊夢"}]},
                {"items": [{"item_id": "changed", "revision": "r2"}]}])) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "apply_edit", **sample_plan(),
                                            "expected_snapshot_hash": editplan.snapshot_hash([])})
        self.assertEqual(result["error_code"], "SNAPSHOT_HASH_MISMATCH")
        self.assertEqual(get.await_count, 2)
        post.assert_not_awaited()

    async def test_reviewed_snapshot_is_rechecked_after_media_and_project_validation(self):
        initial_hash = editplan.snapshot_hash([])
        changed = [video_item()]
        for action in ("apply_edit", "reconcile_edit"):
            gets = [{"characters": []}, {"items": []}, {"success": True, "exists": True},
                    {"success": True, "fps": 30}, {"items": changed}]
            with self.subTest(action=action), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)) as get, \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch({"action": action, **one_video_plan(),
                                                "check_project_settings": True,
                                                "expected_snapshot_hash": initial_hash})
            self.assertEqual(result["error_code"], "SNAPSHOT_HASH_MISMATCH")
            self.assertEqual(result["expected_snapshot_hash"], initial_hash)
            self.assertEqual(result["actual_snapshot_hash"], editplan.snapshot_hash(changed))
            self.assertFalse(result["rolled_back"])
            self.assertEqual(get.await_args.args, ("/items",))
            self.assertEqual(get.await_args_list[-2].args, ("/project",))
            post.assert_not_awaited()

    async def test_reviewed_snapshot_unavailable_after_preflight_stops_apply(self):
        gets = [{"characters": []}, {"items": []}, {"success": True, "exists": True},
                {"success": False, "error_code": "ITEMS_UNAVAILABLE"}]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "apply_edit", **one_video_plan(),
                                            "expected_snapshot_hash": editplan.snapshot_hash([])})
        self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
        self.assertFalse(result["rolled_back"])
        post.assert_not_awaited()

    async def test_matching_snapshot_after_preflight_allows_addition(self):
        item = video_item()
        gets = [{"characters": []}, {"items": []}, {"success": True, "exists": True},
                {"items": []}, {"items": [item]}]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)) as get, \
             patch.object(server, "ymm4_post", AsyncMock(return_value={**item, "success": True})) as post:
            result = await server.dispatch({"action": "apply_edit", **one_video_plan(),
                                            "expected_snapshot_hash": editplan.snapshot_hash([])})
        self.assertTrue(result["success"])
        self.assertEqual(result["added"], 1)
        self.assertEqual(get.await_count, 5)
        post.assert_awaited_once()

    async def test_reviewed_snapshot_rejects_incomplete_or_switched_inventory(self):
        initial = {"items": [], "count": 0, "timeline_id": "timeline:" + "a" * 32}
        for after, code in (
            ({**initial, "count": 1}, "ITEMS_INVENTORY_INCOMPLETE"),
            ({**initial, "timeline_id": "timeline:" + "b" * 32}, "TIMELINE_CHANGED"),
            ({"items": [], "count": 0}, "TIMELINE_ID_UNAVAILABLE"),
        ):
            for action in ("apply_edit", "reconcile_edit"):
                with self.subTest(action=action, code=code):
                    expected = video_item()
                    gets = [{"characters": []}, initial, {"success": True, "exists": True},
                            after, {**initial, "items": [expected], "count": 1}]
                    with (
                        patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)),
                        patch.object(server, "ymm4_post", AsyncMock(return_value={
                            **expected, "success": True})) as post,
                    ):
                        result = await server.dispatch({"action": action, **one_video_plan(),
                                                        "expected_snapshot_hash": editplan.snapshot_hash([])})
                    post.assert_not_awaited()
                    self.assertEqual(result["error_code"], code)
                    self.assertFalse(result["rolled_back"])

    async def test_expected_plan_hash_mismatch_stops_before_host_call(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "apply_edit", **sample_plan(),
                                            "expected_plan_hash": "0" * 64})
        self.assertEqual(result["error_code"], "PLAN_HASH_MISMATCH")
        self.assertEqual(len(result["actual_plan_hash"]), 64)
        get.assert_not_awaited()
        post.assert_not_awaited()

    async def test_actual_voice_length_stops_next_scene_before_overlap(self):
        args = {"action": "apply_edit", "plan": {"scenes": [
            {"id": "intro", "items": [{"id": "voice", "type": "voice",
                                       "character": "A", "text": "hi", "frame": 0}]},
            {"id": "next", "items": [{"id": "title", "type": "text",
                                      "text": "next", "length": 10}]},
        ]}}
        async def get(path):
            if path == "/characters": return {"success": True, "characters": [{"name": "A"}]}
            if path == "/items": return {"success": True, "items": []}
            raise AssertionError(path)
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={
                 "success": True, "item_id": "native:voice", "revision": "r1",
                 "frame": 0, "layer": 0, "length": 100})) as post:
            result = await server.dispatch(args)
        self.assertEqual(result["error_code"], "SCENE_BOUNDARY_CONFLICT")
        self.assertEqual(result["committed_scenes"], ["intro"])
        self.assertEqual((result["actual_end_frame"], result["next_start_frame"]), (100, 30))
        post.assert_awaited_once()

    async def test_project_settings_gate_blocks_wrong_fps_before_edit(self):
        async def get(path):
            if path == "/characters": return {"success": True, "characters": []}
            if path == "/items": return {"success": True, "items": []}
            if path == "/edits/state": return {"success": True, "bindings": []}
            if path == "/project": return {"success": True, "fps": 60}
            raise AssertionError(path)
        args = {**one_text_plan(), "check_project_settings": True}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            preview = await server.dispatch({"action": "plan_edit", **args})
            applied = await server.dispatch({"action": "apply_edit", **args})
        self.assertFalse(preview["passed"])
        self.assertEqual(preview["warnings"][-1]["code"], "PROJECT_SETTINGS_MISMATCH")
        self.assertEqual(applied["error_code"], "PROJECT_SETTINGS_MISMATCH")
        post.assert_not_awaited()

    def _patches(self, items, characters, posts, bindings=None):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": characters}
            if path == "/items":
                return {"success": True, "items": items() if callable(items) else items}
            if path == "/edits/state":
                return {"success": True, "bindings": bindings or []}
            if path.startswith("/media/info?path="):
                return {"success": True, "exists": True}
            raise AssertionError(path)

        return patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
            server, "ymm4_post", AsyncMock(side_effect=posts) if callable(posts) or isinstance(posts, list)
            else AsyncMock(return_value=posts))

    async def test_plan_edit_never_posts(self):
        get, post = self._patches([], [{"name": "ゆっくり霊夢"}], {"success": True})
        with get, post as posted:
            result = await server.dispatch({"action": "plan_edit", **sample_plan()})
            posted.assert_not_awaited()
        self.assertTrue(result["dry_run"])
        self.assertEqual(result["added"], 3)

    async def test_unknown_existing_voice_character_prevents_write_on_apply_and_reconcile(self):
        get, post = self._patches([voice_item()], [{"name": "ゆっくり霊夢"}], {"success": True})
        with get, post as posted:
            for action in ("apply_edit", "reconcile_edit"):
                with self.subTest(action=action):
                    result = await server.dispatch({"action": action, "dry_run": False,
                                                    **one_voice_plan(frame=0)})
                    self.assertEqual(result["error_code"], "EDIT_STATE_CONFLICT")
                    self.assertEqual(result["conflicts"][0]["reason"], "VOICE_CHARACTER_UNAVAILABLE")
            posted.assert_not_awaited()

    async def test_ambiguous_media_prevents_new_writes_in_apply_and_reconcile(self):
        current = [video_item(), video_item("native:duplicate")]
        for action in ("apply_edit", "reconcile_edit"):
            get, post = self._patches(current, [], {"success": True})
            with self.subTest(action=action), get, post as posted:
                result = await server.dispatch({"action": action, **one_video_plan()})
                posted.assert_not_awaited()
            self.assertEqual(result["error_code"], "EDIT_STATE_CONFLICT")
            self.assertEqual(result["conflicts"][0]["candidate_item_ids"],
                             ["native:clip", "native:duplicate"])

    async def test_apply_and_reconcile_preserve_a_later_items_saved_binding(self):
        request = one_text_plan()
        planned = request["plan"]["scenes"][0]["items"]
        planned[0].pop("frame")
        planned.append({**planned[0], "id": "later-caption", "frame": 20})
        bound_later = {"id": "later-caption", "item_id": "native:later", "revision": "r1"}
        record = {"idempotency_key": request["idempotency_key"],
                  "plan_hash": editplan.plan_hash(editplan.parse_plan(request)),
                  "items": [bound_later]}
        for action in ("apply_edit", "reconcile_edit"):
            with self.subTest(action=action):
                current = [{"item_id": "native:later", "revision": "r1", "type": "TextItem",
                            "text": "hello", "frame": 20, "layer": 1, "length": 10}]

                async def post(path, body=None, **unused):
                    if path == "/items/text":
                        added = {"success": True, "item_id": "native:first", "revision": "r2",
                                 **{key: body[key] for key in ("frame", "layer", "length")}}
                        current.append({**added, "type": "TextItem", "text": body["text"]})
                        return added
                    if path == "/edits/bindings":
                        return {"success": True}
                    raise AssertionError(path)

                get, patch_post = self._patches(current, [], post, [record])
                with get, patch_post as posted:
                    result = await server.dispatch({"action": action, **request})
                self.assertTrue(result["success"], result)
                self.assertEqual([(row["id"], row["op"], row["item_id"])
                                  for row in result["details"]],
                                 [("caption", "add", "native:first"),
                                  ("later-caption", "keep", "native:later")])
                self.assertEqual(sorted(item["frame"] for item in current), [0, 20])
                self.assertEqual(posted.await_args.args[1]["items"][1], bound_later)

    async def test_apply_and_reconcile_preserve_later_explicit_media_match_without_bindings(self):
        request = one_video_plan()
        planned = request["plan"]["scenes"][0]["items"]
        planned[0].pop("frame")
        planned.append({**planned[0], "id": "later-clip", "frame": 20})
        for action in ("plan_edit", "apply_edit", "reconcile_edit"):
            with self.subTest(action=action):
                current = [video_item("native:later", frame=20)]

                async def post(path, body=None, **unused):
                    self.assertEqual(path, "/items/video")
                    added = video_item("native:first", frame=body["frame"],
                                       layer=body["layer"], length=body["length"])
                    current.append(added)
                    return {"success": True, **added}

                get, patch_post = self._patches(current, [], post)
                with get, patch_post as posted:
                    result = await server.dispatch({"action": action, **request})
                self.assertTrue(result["success"], result)
                if action == "plan_edit":
                    self.assertTrue(result["passed"], result)
                    operations = result["ops"]
                    posted.assert_not_awaited()
                else:
                    operations = result["details"]
                    self.assertEqual(sorted(item["frame"] for item in current), [0, 20])
                    posted.assert_awaited_once()
                self.assertEqual([(row["id"], row["op"]) for row in operations],
                                 [("clip", "add"), ("later-clip", "keep")])
                self.assertEqual(operations[1]["item_id"], "native:later")

    async def test_explicit_overlay_does_not_move_sequential_frame_before_voice_end(self):
        request = {"action": "apply_edit", "plan": {"gap": 0, "scenes": [{
            "id": "scene", "items": [
                {"id": "voice", "type": "dialogue", "character": "A", "text": "line",
                 "frame": 0, "layer": 1},
                {"id": "overlay", "type": "text", "text": "overlay",
                 "frame": 0, "layer": 2, "length": 10},
                {"id": "tail", "type": "text", "text": "sequential", "layer": 1, "length": 10},
            ]}]}}
        current = []

        async def post(path, body=None, **unused):
            is_voice = path == "/items/voice"
            result = {"success": True, "item_id": f"native:{len(current)}", "revision": "r1",
                      "frame": body["frame"], "layer": body["layer"],
                      "length": 100 if is_voice else body["length"]}
            current.append({**result, "type": "VoiceItem" if is_voice else "TextItem",
                            "text": body["text"]})
            return result

        get, patch_post = self._patches(current, [{"name": "A"}], post)
        with get, patch_post:
            result = await server.dispatch(request)
        self.assertTrue(result["success"], result)
        self.assertEqual([(item["text"], item["frame"]) for item in current],
                         [("line", 0), ("overlay", 0), ("sequential", 100)])

    async def test_apply_edit_adds_then_replays(self):
        created = []
        created_items = []

        async def post(path, body=None, timeout=10.0):
            del timeout
            if path.startswith("/items/"):
                created.append((path, body))
                kind = path.rsplit("/", 1)[-1]
                result = {"success": True, "item_id": f"native:{len(created)}", "revision": "r",
                          "frame": body.get("frame", 0), "layer": body["layer"],
                          "length": 45 if kind == "voice" else body.get("length", 30)}
                created_items.append({**result, "type": {"video": "VideoItem", "voice": "VoiceItem",
                                                          "text": "TextItem"}[kind], "text": body.get("text", body.get("path", ""))})
                return result
            if path == "/edits/bindings":
                return {"success": True, "idempotency_key": body["idempotency_key"]}
            raise AssertionError(path)

        get, posted = self._patches(lambda: created_items, [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", **sample_plan()})
        self.assertTrue(result["success"])
        self.assertEqual(result["added"], 3)
        self.assertEqual([path for path, _ in created], ["/items/video", "/items/voice", "/items/text"])
        self.assertEqual(created[1][1]["character"], "ゆっくり霊夢")

        record = {
            "idempotency_key": "episode-1",
            "plan_hash": result["plan_hash"],
            "items": [{"id": row["id"], "item_id": row["item_id"], "revision": "r"} for row in result["details"]],
        }
        current = created_items

        async def get_replay(path):
            if path == "/characters":
                return {"success": True, "characters": [{"name": "ゆっくり霊夢"}]}
            if path == "/items":
                return {"success": True, "items": current}
            if path == "/edits/state":
                return {"success": True, "bindings": [record]}
            if path.startswith("/media/info?path="):
                return {"success": True, "exists": True}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get_replay)), patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as replay_post:
            replayed = await server.dispatch({"action": "apply_edit", **sample_plan()})
            replay_post.assert_not_awaited()
        self.assertTrue(replayed["replayed"])
        self.assertEqual(replayed["added"], 0)

    async def test_invalid_plan_never_touches_http(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get, patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "plan_edit", "plan": {"scenes": []}})
            get.assert_not_awaited()
            post.assert_not_awaited()

    async def test_missing_host_source_is_reported_before_any_mutation(self):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": [{"name": "ゆっくり霊夢"}]}
            if path == "/items":
                return {"success": True, "items": []}
            if path == "/edits/state":
                return {"success": True, "bindings": []}
            if path.startswith("/media/info?path="):
                return {"success": True, "exists": False}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            preview = await server.dispatch({"action": "plan_edit", **sample_plan()})
            self.assertFalse(preview["passed"])
            self.assertTrue(any(w["code"] == "SOURCE_MISSING" for w in preview["warnings"]))
            applied = await server.dispatch({"action": "apply_edit", **sample_plan()})
            self.assertEqual(applied["error_code"], "SOURCE_MISSING")
            post.assert_not_awaited()

    async def test_unknown_character_is_rejected_before_add(self):
        get, post = self._patches([], [{"name": "ゆっくり魔理沙"}], {"success": True})
        with get, post as posted:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "apply_edit", **sample_plan()})
            posted.assert_not_awaited()

    async def test_partial_failure_rolls_back_the_failed_scene(self):
        calls = []

        async def post(path, body=None, timeout=10.0):
            del timeout
            calls.append((path, body))
            if path == "/items/video":
                return {"success": True, "item_id": "native:bg", "revision": "r", "frame": 0, "length": 300}
            if path == "/items/voice":
                return {"success": False, "error": "character missing"}
            if path == "/items/delete":
                return {"success": True, "removed": 1, "item_ids": body["item_ids"]}
            raise AssertionError(path)

        get, posted = self._patches([], [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", **sample_plan()})
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "EDIT_PARTIAL_FAILURE")
        self.assertTrue(result["rolled_back"])
        self.assertEqual(result["added"], 0)
        self.assertEqual(result["failed_scene"], "intro")
        self.assertEqual(result["committed_scenes"], [])
        self.assertEqual(calls[-1], ("/items/delete", {"item_ids": ["native:bg"]}))

    async def test_unknown_add_outcome_does_not_delete(self):
        calls = []

        async def post(path, body=None, timeout=10.0):
            del timeout, body
            calls.append(path)
            if path == "/items/video":
                raise httpx.ConnectError("connection lost")
            raise AssertionError(path)

        get, posted = self._patches([], [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", **sample_plan()})
        self.assertEqual(result["error_code"], "EDIT_REQUEST_FAILED")
        self.assertTrue(result["outcome_unknown"])
        self.assertFalse(result["rolled_back"])
        self.assertNotIn("/items/delete", calls)

    async def test_atomic_scenes_false_keeps_partial_adds(self):
        async def post(path, body=None, timeout=10.0):
            del timeout
            if path == "/items/video":
                return {"success": True, "item_id": "native:bg", "revision": "r", "frame": 0, "length": 300}
            if path == "/items/voice":
                return {"success": False, "error": "character missing"}
            if path == "/edits/bindings":
                return {"success": True, "idempotency_key": body["idempotency_key"]}
            raise AssertionError(path)

        get, posted = self._patches([], [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", "atomic_scenes": False, **sample_plan()})
        self.assertFalse(result["success"])
        self.assertFalse(result["rolled_back"])
        self.assertEqual(result["added"], 1)
        self.assertEqual(result["details"][0]["item_id"], "native:bg")

    async def test_failed_scene_rolls_back_only_that_scene(self):
        calls = []

        async def post(path, body=None, timeout=10.0):
            del timeout
            calls.append((path, body))
            if path == "/items/text":
                text = body["text"]
                if text == "keep-me":
                    return {"success": True, "item_id": "native:intro", "revision": "r", "frame": 0, "length": 10}
                if text == "temp":
                    return {"success": True, "item_id": "native:body", "revision": "r", "frame": 20, "length": 10}
                return {"success": False, "error": "boom"}
            if path == "/edits/bindings":
                return {"success": True, "idempotency_key": body["idempotency_key"]}
            if path == "/items/delete":
                return {"success": True, "removed": 1, "item_ids": body["item_ids"]}
            raise AssertionError(path)

        plan = {"plan": {"scenes": [
            {"id": "intro", "items": [{"id": "t1", "type": "text", "text": "keep-me", "length": 10, "layer": 1}]},
            {"id": "body", "items": [
                {"id": "t2", "type": "text", "text": "temp", "length": 10, "layer": 1, "frame": 20},
                {"id": "t3", "type": "text", "text": "fail", "length": 10, "layer": 1, "frame": 40},
            ]},
        ]}, "idempotency_key": "ep-2"}
        get, posted = self._patches([], [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", **plan})
        self.assertFalse(result["success"])
        self.assertTrue(result["rolled_back"])
        self.assertEqual(result["committed_scenes"], ["intro"])
        self.assertEqual(result["failed_scene"], "body")
        self.assertEqual(result["added"], 1)
        self.assertEqual(result["details"][0]["item_id"], "native:intro")
        deletes = [body for path, body in calls if path == "/items/delete"]
        self.assertEqual(deletes, [{"item_ids": ["native:body"]}])
        bindings = [body for path, body in calls if path == "/edits/bindings"]
        self.assertEqual(bindings[0]["items"][0]["item_id"], "native:intro")

    async def test_checkpoint_and_rollback_dispatch(self):
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True, "checkpoint_id": "cp_" + "a" * 32})) as post:
            await server.dispatch({"action": "control", "sub_action": "checkpoint", "reason": "before apply", "backup": True})
            post.assert_awaited_once_with("/edits/checkpoint", {"reason": "before apply", "backup": True})
        checkpoint_id = "cp_" + "b" * 32
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            await server.dispatch({"action": "control", "sub_action": "rollback", "checkpoint_id": checkpoint_id})
            post.assert_awaited_once_with("/edits/rollback", {"checkpoint_id": checkpoint_id})
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "checkpoints": []})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "checkpoints"})
            get.assert_awaited_once_with("/edits/checkpoints")
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "checkpoints", "checkpoint_id": checkpoint_id})
            get.assert_awaited_once_with(f"/edits/checkpoints/{checkpoint_id}")
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "rollback", "checkpoint_id": "../x"})
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "checkpoint", "backup": "yes"})
            post.assert_not_awaited()

    async def test_reused_key_with_changed_plan_stops_before_post(self):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"items": []}
            if path == "/edits/state":
                return {"success": True, "bindings": [one_text_record()]}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan("changed")})
            posted.assert_not_awaited()
        self.assertEqual(result["error_code"], "IDEMPOTENCY_KEY_CONFLICT")

    async def test_manual_change_stops_before_post(self):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"items": [{"item_id": "native:caption", "revision": "r2"}]}
            if path == "/edits/state":
                return {"success": True, "bindings": [one_text_record()]}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as posted:
            result = await server.dispatch({"action": "reconcile_edit", **one_text_plan()})
            posted.assert_not_awaited()
        self.assertEqual(result["error_code"], "EDIT_STATE_CONFLICT")

    async def test_missing_item_is_added_once_by_reconcile(self):
        current = []
        posts = []

        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"items": list(current)}
            if path == "/edits/state":
                return {"success": True, "bindings": [one_text_record()]}
            raise AssertionError(path)

        async def post(path, body=None, timeout=10.0):
            del timeout
            posts.append(path)
            if path == "/items/text":
                item = {"item_id": "native:caption", "revision": "r1", "frame": body["frame"],
                        "layer": body["layer"], "length": body["length"], "type": "TextItem", "text": body["text"]}
                current.append(item)
                return {"success": True, **item}
            if path == "/edits/bindings":
                return {"success": True}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", AsyncMock(side_effect=post)):
            result = await server.dispatch({"action": "reconcile_edit", **one_text_plan()})
        self.assertTrue(result["success"])
        self.assertEqual(result["added"], 1)
        self.assertEqual(posts, ["/items/text", "/edits/bindings"])

    async def test_post_apply_mismatch_does_not_save_binding(self):
        items_calls = 0

        async def get(path):
            nonlocal items_calls
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/edits/state":
                return {"success": True, "bindings": []}
            if path == "/items":
                items_calls += 1
                return {"items": [] if items_calls == 1 else [
                    {"item_id": "native:caption", "revision": "r1", "frame": 5, "layer": 1, "length": 10}]}
            raise AssertionError(path)

        async def post(path, body=None, timeout=10.0):
            del body, timeout
            self.assertEqual(path, "/items/text")
            return {"success": True, "item_id": "native:caption", "revision": "r1",
                    "frame": 0, "layer": 1, "length": 10}

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", AsyncMock(side_effect=post)) as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
            self.assertEqual(posted.await_count, 1)
        self.assertEqual(result["error_code"], "EDIT_VERIFY_FAILED")
        self.assertEqual(result["verification_failures"][0]["fields"], ["frame"])

    async def test_retry_after_mismatch_stops_without_adding(self):
        current = [{"item_id": "native:caption", "revision": "r1", "type": "TextItem",
                    "text": "hello", "frame": 5, "layer": 1, "length": 10}]
        get, post = self._patches(current, [], {"success": True})
        with get, post as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
            posted.assert_not_awaited()
        self.assertEqual(result["error_code"], "EDIT_STATE_CONFLICT")

    async def test_post_apply_read_failure_does_not_save_binding(self):
        items_calls = 0

        async def get(path):
            nonlocal items_calls
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/edits/state":
                return {"success": True, "bindings": []}
            if path == "/items":
                items_calls += 1
                if items_calls == 2:
                    raise httpx.ConnectError("connection lost")
                return {"items": []}
            raise AssertionError(path)

        async def post(path, body=None, timeout=10.0):
            del body, timeout
            self.assertEqual(path, "/items/text")
            return {"success": True, "item_id": "native:caption", "revision": "r1",
                    "frame": 0, "layer": 1, "length": 10}

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", AsyncMock(side_effect=post)) as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
            self.assertEqual(posted.await_count, 1)
        self.assertEqual(result["error_code"], "EDIT_VERIFY_FAILED")
        self.assertTrue(result["outcome_unknown"])

    async def test_binding_save_failure_is_not_success(self):
        current = []

        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"items": list(current)}
            if path == "/edits/state":
                return {"success": True, "bindings": []}
            raise AssertionError(path)

        async def post(path, body=None, timeout=10.0):
            del timeout
            if path == "/items/text":
                current.append({"item_id": "native:caption", "revision": "r1", "frame": 0,
                                "layer": 1, "length": 10, "type": "TextItem", "text": "hello"})
                return {"success": True, **current[0]}
            if path == "/edits/bindings":
                return {"success": False, "error_code": "BINDINGS_NOT_SAVED"}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", AsyncMock(side_effect=post)):
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "BINDINGS_NOT_SAVED")
        self.assertEqual(result["added"], 1)
        self.assertFalse(result["rolled_back"])

    async def test_unavailable_bindings_stop_before_post(self):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"items": []}
            if path == "/edits/state":
                return {"success": False, "error": "disk error"}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
            posted.assert_not_awaited()
        self.assertEqual(result["error_code"], "BINDINGS_UNAVAILABLE")

    async def test_invalid_initial_items_stop_before_post(self):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": []}
            if path == "/items":
                return {"success": True, "items": None}
            raise AssertionError(path)

        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), patch.object(
                server, "ymm4_post", new_callable=AsyncMock) as posted:
            result = await server.dispatch({"action": "apply_edit", **one_text_plan()})
            posted.assert_not_awaited()
        self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")

    async def test_get_edit_state_and_skill_prompt(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "scenes": []})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "edit_state"})
            get.assert_awaited_once_with("/edits/state")
        import mcp_skills
        prompt = await mcp_skills.get_prompt("jikkyou", {"theme": "ボス戦"})
        self.assertIn("plan_edit", prompt.messages[0].content.text)
        self.assertIn("reconcile_edit", prompt.messages[0].content.text)
        self.assertIn("control/checkpoint", prompt.messages[0].content.text)
        self.assertIn("control/rollback", prompt.messages[0].content.text)


if __name__ == "__main__":
    unittest.main()
