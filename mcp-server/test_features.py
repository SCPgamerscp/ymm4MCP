"""Regression tests for authenticated discovery, skill exposure and safe editing."""
import asyncio
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import AsyncMock, patch

import httpx
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

import editing
import mcp_skills
import server
from ymm4_connection import connection_settings, ConnectionConfigurationError

ROOT = Path(__file__).resolve().parent


class ConnectionTests(unittest.TestCase):
    def test_missing_credentials_fail_closed(self):
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(ConnectionConfigurationError):
                connection_settings()

    def test_environment_token_and_port(self):
        with patch.dict(os.environ, {"YMM4_API_TOKEN": "test-token", "YMM4_API_BASE": "http://127.0.0.1:9876/api/"}, clear=True):
            self.assertEqual(connection_settings(), ("http://127.0.0.1:9876/api", {"X-Ymm4-Token": "test-token"}))

    def test_remote_and_malformed_urls_rejected(self):
        for base in ("http://example.com/api", "https://127.0.0.1/api", "http://127.0.0.1.evil/api",
                     "http://user@localhost/api", "http://localhost/api?secret=x", "http://localhost:80/api",
                     "http://localhost:99999/api", "http://localhost/other", "http://localhost/api#fragment"):
            with self.subTest(base=base), patch.dict(os.environ, {"YMM4_API_TOKEN": "secret", "YMM4_API_BASE": base}, clear=True):
                with self.assertRaises(ValueError):
                    connection_settings()

    def test_descriptor_is_reread_after_rotation(self):
        with tempfile.TemporaryDirectory(dir=ROOT) as directory:
            path = Path(directory) / "connection.json"
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": str(path)}, clear=True):
                for token, port in (("old-token", 8765), ("new-token", 9876)):
                    path.write_text(json.dumps({"token": token, "api_base": f"http://127.0.0.1:{port}/api"}))
                    base, headers = connection_settings()
                    self.assertIn(str(port), base)
                    self.assertEqual(headers["X-Ymm4-Token"], token)
                for content in ("not json", "[]", '{}', '{"token": "bad\\nvalue"}'):
                    path.write_text(content)
                    with self.assertRaises(ConnectionConfigurationError):
                        connection_settings()

    def test_environment_token_does_not_need_descriptor(self):
        with patch.dict(os.environ, {"YMM4_API_TOKEN": "secret", "YMM4_CONNECTION_FILE": "does-not-exist"}, clear=True):
            self.assertEqual(connection_settings()[1]["X-Ymm4-Token"], "secret")


class PlanningTests(unittest.TestCase):
    def test_qa_gate_passes_or_repairs_within_budget(self):
        failed = editing.validate_timeline([
            {"item_id": "a", "frame": 0, "layer": 0, "length": 20},
            {"item_id": "b", "frame": 10, "layer": 0, "length": 20},
        ])
        result = editing.evaluate_qa_gate(failed)
        self.assertEqual((result["decision"], result["reason_code"]), ("repair", "QA_ISSUES_REMAIN"))
        passed = editing.validate_timeline([{"item_id": "a", "frame": 0, "layer": 0, "length": 20}])
        result = editing.evaluate_qa_gate(passed, [failed], max_repairs=1)
        self.assertEqual(result["decision"], "pass")
        self.assertEqual(result["score_delta"], 20)

    def test_qa_gate_stops_on_regression_stall_and_limits(self):
        clean = editing.validate_timeline([{"frame": 0, "layer": 0, "length": 20}])
        failed = editing.validate_timeline([
            {"frame": 0, "layer": 0, "length": 20},
            {"frame": 10, "layer": 0, "length": 20},
        ])
        result = editing.evaluate_qa_gate(failed, [clean])
        self.assertEqual((result["reason_code"], result["suggested_action"]),
                         ("QA_REGRESSED", "consider_checkpoint_rollback"))
        self.assertEqual(editing.evaluate_qa_gate(failed, [failed])["reason_code"], "QA_STALLED")
        changed = {**failed, "issues": [{"code": "OTHER"}]}
        self.assertEqual(editing.evaluate_qa_gate(changed, [failed], max_repairs=1)["reason_code"],
                         "REPAIR_LIMIT_REACHED")
        self.assertEqual(editing.evaluate_qa_gate(failed, max_seconds=30, elapsed_seconds=30)["reason_code"],
                         "TIME_LIMIT_REACHED")
        self.assertEqual(editing.evaluate_qa_gate(failed, max_api_calls=5, api_calls=5)["reason_code"],
                         "API_LIMIT_REACHED")

    def test_qa_gate_rejects_invalid_history_and_limits(self):
        good = editing.validate_timeline([])
        for kwargs in ({"history": [{}]}, {"history": [good] * 21}, {"max_repairs": True},
                       {"elapsed_seconds": float("nan")}, {"max_seconds": "no"},
                       {"api_calls": -1}, {"repeat_limit": 1}):
            with self.subTest(kwargs=kwargs), self.assertRaises(ValueError):
                editing.evaluate_qa_gate(good, **kwargs)

    def test_invalid_script_rejected_before_execution(self):
        base = {"lines": [{"character": "霊夢", "text": "説明"}]}
        invalid = [{"chars_per_sec": 0}, {"chars_per_sec": float("nan")}, {"chars_per_sec": float("inf")},
                   {"chars_per_sec": 1e-320}, {"fps": True}, {"fps": 0}, {"start_frame": -1}, {"gap": -1},
                   {"lines": []}, {"lines": [None]}, {"lines": [{"character": "霊夢", "text": " "}]},
                   {"lines": [{"character": "霊夢", "text": "hi", "layer": True}]}, {"dry_run": "false"}]
        for change in invalid:
            with self.subTest(change=change), self.assertRaises(ValueError):
                editing.plan_script({**base, **change})

    def test_auto_layers_avoid_explicit_layers_and_estimates_are_labelled(self):
        result = editing.plan_script({"start_frame": 100, "gap": 5, "lines": [
            {"character": "A", "text": "Hi", "layer": 0}, {"character": "B", "text": "Hello"},
            {"character": "B", "text": "Again"}, {"character": "C", "text": "Last"},
        ]})
        self.assertEqual([x["layer"] for x in result["details"]], [0, 1, 1, 2])
        self.assertEqual(result["total_frames"], 240)
        self.assertTrue(result["estimated"])
        self.assertEqual(result["added"], 0)

    def test_finite_number_accepts_numeric_strings_and_rejects_non_finite(self):
        self.assertEqual(editing.finite_number("1.5", "value"), 1.5)
        self.assertEqual(editing.finite_number(0, "value"), 0.0)
        for bad in (True, "nope", float("nan"), float("inf"), None, []):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                editing.finite_number(bad, "value")

    def test_validation_returns_structured_qa_with_stable_ids(self):
        items = [
            {"item_id": "native:a", "revision": "r1", "frame": 10, "layer": 0, "length": 100},
            {"item_id": "native:b", "revision": "r2", "frame": 15, "layer": 0, "length": 10},
            {"item_id": "native:c", "revision": "r3", "frame": 30, "layer": 0, "length": 10},
            {"item_id": "native:d", "revision": "r4", "frame": 120, "layer": 0, "length": 20},
        ]
        result = editing.validate_timeline(items, expected=[{"item_id": "missing"}], duration=130)
        codes = [issue["code"] for issue in result["issues"]]
        self.assertEqual(codes.count("OVERLAP"), 2)
        self.assertIn("GAP", codes)
        self.assertIn("EXCEEDS_DURATION", codes)
        self.assertIn("EXPECTED_NOT_FOUND", codes)
        overlap = next(issue for issue in result["issues"] if issue["code"] == "OVERLAP")
        self.assertEqual(overlap["item_ids"], ["native:a", "native:b"])
        self.assertEqual(overlap["frame_range"], [15, 25])
        self.assertEqual(result["summary"], {"errors": 4, "warnings": 1})
        self.assertEqual(result["issue_count"], 5)
        self.assertEqual(result["problems"], result["issues"])
        self.assertFalse(result["passed"])
        self.assertFalse(result["valid"])

    def test_different_layers_touching_edges_and_leading_space_are_valid(self):
        items = [{"frame": 10, "layer": 0, "length": 10}, {"frame": 20, "layer": 0, "length": 10},
                 {"frame": 10, "layer": 1, "length": 20}]
        result = editing.validate_timeline(items, [{"frame": 20, "layer": 0}], duration=30)
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"], [])
        self.assertEqual(result["score"], 100)

    def test_gap_reporting_can_be_disabled(self):
        items = [{"frame": 0, "layer": 0, "length": 10}, {"frame": 20, "layer": 0, "length": 10}]
        result = editing.validate_timeline(items, include_gaps=False)
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"], [])
        with self.assertRaisesRegex(ValueError, "include_gaps"):
            editing.validate_timeline(items, include_gaps="false")

    def test_subtitle_qa_matches_text_and_time_only_on_selected_layers(self):
        items = [
            {"item_id": "voice:1", "type": "VoiceItem", "text": "こんにちは 世界", "frame": 10, "layer": 7, "length": 60},
            {"item_id": "caption:1", "type": "TextItem", "text": "こんにちは\n世界", "frame": 12, "layer": 9, "length": 58},
            {"item_id": "title", "type": "TextItem", "text": "別のテロップ", "frame": 10, "layer": 3, "length": 60},
        ]
        result = editing.validate_timeline(items, include_gaps=False, subtitle_layers=[9])
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"], [])
        self.assertTrue(editing.validate_timeline(items, include_gaps=False)["passed"])

    def test_subtitle_qa_reports_missing_mismatch_and_unreadable_voice(self):
        items = [
            {"item_id": "voice:1", "type": "VoiceItem", "text": "正しいセリフ", "frame": 0, "layer": 7, "length": 50},
            {"item_id": "caption:wrong", "type": "TextItem", "text": "別の字幕", "frame": 0, "layer": 9, "length": 50},
            {"item_id": "voice:2", "type": "VoiceItem", "text": "次のセリフ", "frame": 50, "layer": 7, "length": 50},
            {"item_id": "voice:3", "type": "VoiceItem", "text": None, "frame": 100, "layer": 7, "length": 50},
            {"item_id": "title", "type": "TextItem", "text": "次のセリフ", "frame": 50, "layer": 3, "length": 50},
        ]
        result = editing.validate_timeline(items, include_gaps=False, subtitle_layers=[9])
        self.assertEqual([issue["code"] for issue in result["issues"]], [
            "SUBTITLE_TEXT_MISMATCH", "SUBTITLE_MISSING", "SUBTITLE_CHECK_SKIPPED"])
        mismatch = result["issues"][0]
        self.assertEqual(mismatch["item_ids"], ["voice:1", "caption:wrong"])
        self.assertEqual(mismatch["actual_texts"], ["別の字幕"])
        self.assertEqual(result["issues"][1]["frame_range"], [50, 100])
        self.assertEqual(result["summary"], {"errors": 2, "warnings": 1})
        self.assertFalse(result["passed"])
        self.assertTrue(editing.validate_timeline(items, include_gaps=False)["passed"])

    def test_subtitle_layers_must_be_nonempty_integer_list(self):
        for layers in ([], [True], [-1], "9", [0] * 129):
            with self.subTest(layers=layers), self.assertRaisesRegex(ValueError, "subtitle_layers"):
                editing.validate_timeline([], subtitle_layers=layers)

    def test_expected_stable_id_revision_and_legacy_id(self):
        item = {"item_id": "native:one", "revision": "abc", "frame": 0, "layer": 0, "length": 10}
        for expected in ([{"item_id": "native:one", "revision": "abc"}], [{"id": "native:one"}]):
            with self.subTest(expected=expected):
                self.assertTrue(editing.validate_timeline([item], expected)["passed"])
        with self.assertRaisesRegex(ValueError, "must not conflict"):
            editing.validate_timeline([item], [{"id": "one", "item_id": "two"}])

    def test_ambiguous_invalid_and_overflow_items(self):
        item = {"item_id": "runtime:a", "frame": 0, "layer": 0, "length": 10}
        result = editing.validate_timeline([item, item, {"item_id": "runtime:b", "frame": 2147483647, "layer": 2, "length": 1}], [item])
        self.assertIn("EXPECTED_AMBIGUOUS", [p["code"] for p in result["problems"]])
        invalid = next(p for p in result["issues"] if p["code"] == "INVALID_ITEM")
        self.assertEqual(invalid["item_ids"], ["runtime:b"])
        with self.assertRaises(ValueError):
            editing.validate_timeline([], [{}])


class FeatureTests(unittest.IsolatedAsyncioTestCase):
    async def test_dry_run_sends_no_http(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get, patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.add_script({"dry_run": True, "lines": [{"character": "霊夢", "text": "説明"}]})
            self.assertTrue(result["dry_run"])
            get.assert_not_awaited()
            post.assert_not_awaited()

    async def test_script_preflight_and_actual_duration(self):
        lines = [{"character": "A", "text": "one"}, {"character": "A", "text": "two"}]
        added = [{"success": True, "item_id": "native:one", "revision": "r1", "frame": 10, "layer": 0, "length": 95},
                 {"success": True, "item_id": "native:two", "revision": "r2", "frame": 110, "layer": 0, "length": 10}]
        snapshot = {"items": [{k: v for k, v in result.items() if k != "success"} for result in added]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
            {"characters": [{"name": "A"}]}, snapshot,
        ])) as get, patch.object(server, "ymm4_post", AsyncMock(side_effect=added)) as post:
            result = await server.add_script({"start_frame": 10, "gap": 5, "lines": lines})
            self.assertEqual(post.await_args_list[1].args[1]["frame"], 110)
            self.assertEqual(result["total_frames"], 125)
            self.assertEqual(result["added"], 2)
            self.assertTrue(result["verified"])
            self.assertEqual(get.await_args_list[-1].args, ("/items",))
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"characters": []})), patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.add_script({"lines": lines})
            post.assert_not_awaited()

    async def test_script_verification_rejects_missing_or_changed_items(self):
        added = {"success": True, "item_id": "native:one", "revision": "r1",
                 "frame": 0, "layer": 0, "length": 30}
        cases = [
            ({"items": []}, "ITEM_NOT_UNIQUE"),
            ({"items": [{**added, "length": 40}]}, "STATE_MISMATCH"),
        ]
        for snapshot, reason in cases:
            with self.subTest(reason=reason), patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"characters": [{"name": "A"}]}, snapshot,
            ])), patch.object(server, "ymm4_post", AsyncMock(return_value=added)) as post:
                result = await server.add_script({"lines": [{"character": "A", "text": "one"}]})
                self.assertEqual(result["error_code"], "SCRIPT_VERIFY_FAILED")
                self.assertEqual(result["verification_failures"][0]["reason"], reason)
                self.assertTrue(result["outcome_unknown"])
                post.assert_awaited_once()

    async def test_script_verification_unavailable_does_not_retry_voice(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
            {"characters": [{"name": "A"}]}, httpx.ReadTimeout(""),
        ])), patch.object(server, "ymm4_post", AsyncMock(return_value={
            "success": True, "item_id": "native:one", "revision": "r1",
            "frame": 0, "layer": 0, "length": 30,
        })) as post:
            result = await server.add_script({"lines": [{"character": "A", "text": "one"}]})
            self.assertEqual(result["error_code"], "SCRIPT_VERIFY_FAILED")
            self.assertTrue(result["outcome_unknown"])
            post.assert_awaited_once()

    async def test_partial_failure_and_unknown_length_stop(self):
        for failure in ({"success": False, "error": "failed"}, {"success": True, "frame": 10, "length": -1}, httpx.ReadTimeout("")):
            with self.subTest(failure=failure), patch.object(server, "ymm4_get", AsyncMock(return_value={"characters": [{"name": "A"}]})), patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                {"success": True, "frame": 0, "length": 10}, failure,
            ])) as post:
                result = await server.add_script({"lines": [{"character": "A", "text": "x"}] * 3})
                self.assertFalse(result["success"])
                self.assertEqual(result["failed_line"], 1)
                self.assertFalse(result["rolled_back"])
                self.assertEqual(post.await_count, 2)

    async def test_invalid_late_line_causes_no_partial_edit(self):
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.add_script({"lines": [{"character": "A", "text": "x"}, {"character": "A", "text": ""}]})
            post.assert_not_awaited()

    async def test_media_dispatch_preserves_paths_and_lengths(self):
        for kind in ("video", "audio", "image"):
            with self.subTest(kind=kind), patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
                await server.dispatch({"action": "add_item", "sub_action": kind, "path": "C:/動画素材/test.mp4", "frame": 10, "layer": 2, "length": 90})
                post.assert_awaited_once_with(f"/items/{kind}", {"path": "C:/動画素材/test.mp4", "frame": 10, "layer": 2, "length": 90}, timeout=120.0)
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "add_item", "sub_action": "image", "path": "C:/a.png"})
            post.assert_not_awaited()

    async def test_media_info_uses_ymm4_host_and_validates_path(self):
        path = "C:/動画素材/clip 01.mp4"
        result = {"success": True, "exists": True, "path": path, "extension": ".mp4", "bytes": 1234}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=result)) as get:
            self.assertEqual(await server.dispatch({"action": "get_info", "sub_action": "media", "path": path}), result)
            get.assert_awaited_once_with("/media/info?path=C%3A%2F%E5%8B%95%E7%94%BB%E7%B4%A0%E6%9D%90%2Fclip%2001.mp4")
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            for invalid in (None, "clip.mp4", "https://example.com/a.mp4", "C:/x\n.mp4"):
                with self.subTest(path=invalid), self.assertRaisesRegex(ValueError, "absolute"):
                    await server.dispatch({"action": "get_info", "sub_action": "media", "path": invalid})
            get.assert_not_awaited()

    async def test_identity_and_revision_are_forwarded_for_safe_edits(self):
        identity = "native:item/with spaces"
        revision = "abc123"
        cases = [
            ("property", "/items/prop", {"frame": 0, "layer": 0, "prop": "Length", "value": "90", "item_id": identity, "expected_revision": revision}),
            ("delete", "/items/delete", {"item_id": identity, "expected_revision": revision}),
            ("select", "/items/select", {"item_id": identity}),
            ("keyframe", "/items/keyframe", {"prop": "X", "action": "set", "item_id": identity,
                                             "expected_revision": revision, "at": 30, "value": -200.0}),
        ]
        for sub_action, path, expected in cases:
            args = {"action": "edit_item", "sub_action": sub_action, "item_id": identity}
            if sub_action == "property": args.update(prop="Length", value=90)
            if sub_action == "keyframe": args.update(prop="X", value=-200, at=30, keyframe_action="set")
            if sub_action != "select": args["expected_revision"] = revision
            with self.subTest(sub_action=sub_action), patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
                await server.dispatch(args)
                post.assert_awaited_once_with(path, expected)

        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": []})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "effects", "item_id": identity})
            get.assert_awaited_once_with("/items/effects?item_id=native%3Aitem%2Fwith%20spaces")
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "keyframes", "item_id": identity, "prop": "X"})
            get.assert_awaited_once_with("/items/keyframes?item_id=native%3Aitem%2Fwith%20spaces&prop=X")

    async def test_revision_requires_item_id_before_http_request(self):
        for sub_action in ("property", "delete", "keyframe"):
            args = {"action": "edit_item", "sub_action": sub_action, "expected_revision": "stale"}
            if sub_action == "property":
                args.update(prop="Length", value=90, frame=10, layer=2)
            if sub_action == "keyframe":
                args.update(prop="X", value=1, at=0)
            with self.subTest(sub_action=sub_action), patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                with self.assertRaisesRegex(ValueError, "item_id"):
                    await server.dispatch(args)
                post.assert_not_awaited()

    async def test_keyframe_set_requires_value_and_rejects_non_finite(self):
        base = {"action": "edit_item", "sub_action": "keyframe", "item_id": "native:a", "prop": "X", "at": 10}
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaisesRegex(ValueError, "value"):
                await server.dispatch(base)
            post.assert_not_awaited()
            with self.assertRaisesRegex(ValueError, "value"):
                await server.dispatch({**base, "value": float("nan")})
            post.assert_not_awaited()
            await server.dispatch({**base, "value": "-12.5", "keyframe_action": "set"})
            post.assert_awaited_once_with("/items/keyframe", {
                "prop": "X", "action": "set", "item_id": "native:a", "at": 10, "value": -12.5,
            })

    async def test_invalid_edit_coordinates_are_rejected_before_http(self):
        invalid = [
            {"sub_action": "property", "frame": -1, "layer": 0, "prop": "Length", "value": 1},
            {"sub_action": "select", "frame": 0, "layer": -1},
            {"sub_action": "move", "filename": "clip.mp4", "frame": -1},
            {"sub_action": "move", "filename": "clip.mp4", "length": 0},
            {"sub_action": "move", "filename": ""},
            {"sub_action": "move", "filename": "  "},
            {"sub_action": "move", "filename": 42},
            {"sub_action": "duration", "frames": 0},
            {"sub_action": "resolve_overlaps", "gap": -1},
            {"sub_action": "resolve_overlaps", "layers": [0, -1]},
            {"sub_action": "shift", "from_frame": -1, "delta": 1},
            {"sub_action": "shift", "from_frame": 0, "delta": True},
        ]
        for values in invalid:
            with self.subTest(values=values), patch.object(
                server, "ymm4_post", new_callable=AsyncMock
            ) as post:
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "edit_item", **values})
                post.assert_not_awaited()

    async def test_delete_keeps_legacy_minus_one_selector_sentinel(self):
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            await server.dispatch({
                "action": "edit_item", "sub_action": "delete", "frame": -1, "layer": -1,
            })
            post.assert_awaited_once_with("/items/delete", {})

    async def test_advanced_hidden_and_rejected_unless_enabled(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertNotIn("ymm4_advanced", [t.name for t in (await server.list_tools()).tools])
            result = await server.call_tool("ymm4_advanced", {"action": "get"})
            self.assertTrue(result.isError)
        with patch.dict(os.environ, {"YMM4_ENABLE_ADVANCED": "1"}):
            self.assertIn("ymm4_advanced", [t.name for t in (await server.list_tools()).tools])

    async def test_advanced_inspect_encodes_query_values(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True})) as get:
            await server.dispatch_advanced({
                "action": "inspect", "target": "Main&debug=1",
                "path": "Items[0].Name&x=1/日本語",
            })
            get.assert_awaited_once_with(
                "/reflect/inspect?target=Main%26debug%3D1&path=Items%5B0%5D.Name%26x%3D1%2F%E6%97%A5%E6%9C%AC%E8%AA%9E"
            )

    async def test_failure_response_is_an_mcp_error(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": False, "error_code": "NO_TIMELINE", "error": "No timeline"})):
            result = await server.call_tool("ymm4_interact", {"action": "get_info", "sub_action": "characters"})
            self.assertTrue(result.isError)
            result = await server.dispatch({"action": "validate"})
            self.assertFalse(result["success"])

    async def test_validate_forwards_gap_option_and_stable_expectations(self):
        snapshot = {"items": [{"item_id": "native:a", "revision": "r1", "frame": 10, "layer": 0, "length": 5}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)) as get:
            result = await server.dispatch({"action": "validate", "include_gaps": False,
                                            "expected": [{"item_id": "native:a", "revision": "r1"}]})
        get.assert_awaited_once_with("/items")
        self.assertTrue(result["passed"])

    async def test_qa_gate_validates_current_snapshot_without_mutation(self):
        snapshot = {"items": [
            {"item_id": "native:a", "frame": 0, "layer": 0, "length": 20},
            {"item_id": "native:b", "frame": 10, "layer": 0, "length": 20},
        ]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "qa_gate", "max_repairs": 0})
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()
        self.assertEqual(result["decision"], "stop")
        self.assertEqual(result["reason_code"], "REPAIR_LIMIT_REACHED")
        self.assertEqual(result["qa"]["issues"][0]["code"], "OVERLAP")

    async def test_validate_forwards_subtitle_layers(self):
        snapshot = {"items": [
            {"item_id": "voice:1", "type": "VoiceItem", "text": "はい", "frame": 0, "layer": 7, "length": 30},
            {"item_id": "caption:1", "type": "TextItem", "text": "はい", "frame": 0, "layer": 9, "length": 30},
        ]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)):
            self.assertTrue((await server.dispatch({"action": "validate", "subtitle_layers": [9]}))["passed"])
            result = await server.dispatch({"action": "validate", "subtitle_layers": [8]})
            self.assertEqual(result["issues"][0]["code"], "SUBTITLE_MISSING")

    async def test_skills_are_allowlisted_and_roles_consistent(self):
        for name in mcp_skills.SKILLS:
            self.assertTrue(await mcp_skills.read_resource(f"ymm4://skills/{name}"))
            prompt = await mcp_skills.get_prompt(name, {"theme": "宇宙", "duration_seconds": "60"})
            self.assertIn("宇宙", prompt.messages[0].content.text)
        for uri in ("ymm4://skills/../config.json", "file:///etc/passwd", "ymm4://skills/unknown"):
            with self.assertRaises(ValueError):
                await mcp_skills.read_resource(uri)
        for args in ({}, {"theme": "x", "duration_seconds": "-1"}, {"theme": "x", "extra": "x"}):
            with self.assertRaises(ValueError):
                await mcp_skills.get_prompt("kaisetsu", args)
        text = mcp_skills.skill_text("kaisetsu")
        self.assertIn("魔理沙が基礎を説明", text)
        self.assertNotIn("霊夢が説明", text)
        self.assertNotIn("魔理沙リアクション", text)

    async def test_stdio_resources_prompts_and_dry_run(self):
        env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", YMM4_ENABLE_ADVANCED="0")
        params = StdioServerParameters(command=sys.executable, args=[str(ROOT / "server.py")], env=env)
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as client:
                await client.initialize()
                self.assertEqual(len((await client.list_resources()).resources), 4)
                self.assertEqual(len((await client.list_prompts()).prompts), 4)
                resource = await client.read_resource("ymm4://skills/kaisetsu")
                self.assertIn("魔理沙", resource.contents[0].text)
                prompt = await client.get_prompt("kaisetsu", {"theme": "テスト"})
                self.assertIn("テスト", prompt.messages[0].content.text)
                names = [tool.name for tool in (await client.list_tools()).tools]
                self.assertNotIn("ymm4_advanced", names)
                result = await client.call_tool("ymm4_interact", {"action": "add_script", "dry_run": True, "lines": [{"character": "A", "text": "hello"}]})
                self.assertFalse(result.isError)
                self.assertEqual(json.loads(result.content[0].text)["added"], 0)


if __name__ == "__main__":
    unittest.main()
