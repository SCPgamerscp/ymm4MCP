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


class ParsePlanTests(unittest.TestCase):
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
        ]
        for args in invalid:
            with self.subTest(args=args), self.assertRaises(ValueError):
                editplan.parse_plan(args)

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

    def test_added_item_verification_checks_identity_and_revision(self):
        detail = {"id": "caption", "op": "add", "item_id": "native:caption",
                  "result": {"revision": "r1", "frame": 0, "layer": 1, "length": 10}}
        self.assertEqual(editplan.verify_applied_items([detail], []), [
            {"id": "caption", "item_id": "native:caption", "reason": "ITEM_NOT_UNIQUE"}])
        current = [{"item_id": "native:caption", "revision": "r2", "frame": 0,
                    "layer": 1, "length": 10}]
        self.assertEqual(editplan.verify_applied_items([detail], current)[0]["fields"], ["revision"])


class EditPlanDispatchTests(unittest.IsolatedAsyncioTestCase):
    def _patches(self, items, characters, posts):
        async def get(path):
            if path == "/characters":
                return {"success": True, "characters": characters}
            if path == "/items":
                return {"success": True, "items": items() if callable(items) else items}
            if path == "/edits/state":
                return {"success": True, "bindings": []}
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

    async def test_unknown_character_is_rejected_before_add(self):
        get, post = self._patches([], [{"name": "ゆっくり魔理沙"}], {"success": True})
        with get, post as posted:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "apply_edit", **sample_plan()})
            posted.assert_not_awaited()

    async def test_partial_failure_does_not_continue(self):
        async def post(path, body=None, timeout=10.0):
            del body, timeout
            if path == "/items/video":
                return {"success": True, "item_id": "native:bg", "revision": "r", "frame": 0, "length": 300}
            if path == "/items/voice":
                return {"success": False, "error": "character missing"}
            raise AssertionError(path)

        get, posted = self._patches([], [{"name": "ゆっくり霊夢"}], post)
        with get, posted:
            result = await server.dispatch({"action": "apply_edit", **sample_plan()})
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "EDIT_PARTIAL_FAILURE")
        self.assertEqual(result["added"], 1)
        self.assertFalse(result["rolled_back"])

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


if __name__ == "__main__":
    unittest.main()
