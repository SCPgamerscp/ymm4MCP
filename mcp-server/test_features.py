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
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"characters": [{"name": "A"}]})), patch.object(server, "ymm4_post", AsyncMock(side_effect=[
            {"success": True, "frame": 10, "length": 95}, {"success": True, "frame": 110, "length": 10}
        ])) as post:
            result = await server.add_script({"start_frame": 10, "gap": 5, "lines": lines})
            self.assertEqual(post.await_args_list[1].args[1]["frame"], 110)
            self.assertEqual(result["total_frames"], 125)
            self.assertEqual(result["added"], 2)
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"characters": []})), patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.add_script({"lines": lines})
            post.assert_not_awaited()

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

    async def test_identity_and_revision_are_forwarded_for_safe_edits(self):
        identity = "native:item/with spaces"
        revision = "abc123"
        cases = [
            ("property", "/items/prop", {"frame": 0, "layer": 0, "prop": "Length", "value": "90", "item_id": identity, "expected_revision": revision}),
            ("delete", "/items/delete", {"item_id": identity, "expected_revision": revision}),
            ("select", "/items/select", {"item_id": identity}),
        ]
        for sub_action, path, expected in cases:
            args = {"action": "edit_item", "sub_action": sub_action, "item_id": identity}
            if sub_action == "property": args.update(prop="Length", value=90)
            if sub_action != "select": args["expected_revision"] = revision
            with self.subTest(sub_action=sub_action), patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
                await server.dispatch(args)
                post.assert_awaited_once_with(path, expected)

        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": []})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "effects", "item_id": identity})
            get.assert_awaited_once_with("/items/effects?item_id=native%3Aitem%2Fwith%20spaces")

    async def test_advanced_hidden_and_rejected_unless_enabled(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertNotIn("ymm4_advanced", [t.name for t in (await server.list_tools()).tools])
            result = await server.call_tool("ymm4_advanced", {"action": "get"})
            self.assertTrue(result.isError)
        with patch.dict(os.environ, {"YMM4_ENABLE_ADVANCED": "1"}):
            self.assertIn("ymm4_advanced", [t.name for t in (await server.list_tools()).tools])

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
