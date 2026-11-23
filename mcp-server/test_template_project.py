"""Template creation must guard source and destination before opening YMM4 projects."""
import unittest
from unittest.mock import AsyncMock, patch
from urllib.parse import quote

import server


ARGS = {"action": "create_from_template", "template_path": "C:/projects/template.ymmp",
        "path": "C:/projects/new.ymmp"}
CAPABILITIES = {"success": True, "features": {"guarded_template_save": True}}


class TemplateProjectTests(unittest.IsolatedAsyncioTestCase):
    async def test_dry_run_only_reads_host_state(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "exists": True}, {"success": True, "exists": False},
                {"success": True, "isSaved": True, "projectPath": "C:/projects/old.ymmp"}])) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({**ARGS, "dry_run": True})
        self.assertTrue(result["dry_run"])
        self.assertEqual(get.await_count, 3)
        post.assert_not_awaited()

    async def test_open_save_as_then_verify(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "exists": True}, {"success": True, "exists": False},
                {"success": True, "isSaved": True}, CAPABILITIES,
                {"success": True, "isSaved": True, "projectPath": "C:/projects/template.ymmp"},
                {"success": True, "isSaved": True, "projectPath": "c:\\projects\\new.ymmp"},
                {"success": True, "exists": True, "bytes": 123, "path": "c:\\projects\\new.ymmp"}])) as get, \
             patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            result = await server.dispatch(ARGS)
        self.assertTrue(result["verified"])
        self.assertEqual([call.args[0] for call in post.await_args_list],
                         ["/project/open", "/project/save-template-as"])
        self.assertEqual(post.await_args_list[-1].args[1], {
            "path": ARGS["path"], "expected_template_path": ARGS["template_path"]})
        self.assertEqual(get.await_args.args, (f"/media/info?path={quote(ARGS['path'], safe='')}",))

    async def test_saved_ui_state_requires_nonempty_file_at_the_destination(self):
        found = {"success": True, "exists": True, "bytes": 123, "path": ARGS["path"]}
        for receipt in ({"success": True, "exists": False, "path": ARGS["path"]},
                        {**found, "bytes": 0}, {**found, "bytes": True},
                        {**found, "path": ARGS["template_path"]},
                        {**found, "error_code": "FILE_UNAVAILABLE"}, None,
                        ValueError("Unreadable media response")):
            with self.subTest(receipt=receipt), patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "exists": True}, {"success": True, "exists": False},
                {"success": True, "isSaved": True}, CAPABILITIES,
                {"success": True, "isSaved": True, "projectPath": ARGS["template_path"]},
                {"success": True, "isSaved": True, "projectPath": ARGS["path"]}, receipt,
            ])) as get, patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
                result = await server.dispatch(ARGS)
                self.assertFalse(result["success"])
                expected = ("TEMPLATE_VERIFY_OUTCOME_UNKNOWN" if isinstance(receipt, ValueError)
                            else "TEMPLATE_VERIFY_FAILED")
                self.assertEqual(result["error_code"], expected)
                self.assertEqual(result["check"], "destination_file")
                self.assertTrue(result["outcome_unknown"])
                self.assertEqual(result["destination_path"], ARGS["path"])
                self.assertNotIn("verified", result)
                self.assertEqual(get.await_count, 7)
                self.assertEqual(post.await_count, 2)

    async def test_existing_destination_and_unsaved_project_do_not_open(self):
        for results, expected in (
            ([{"success": True, "exists": True}, {"success": True, "exists": True}],
             "TEMPLATE_DESTINATION_UNAVAILABLE"),
            ([{"success": True, "exists": True}, {"success": True, "exists": False},
              {"success": True, "isSaved": False}], "CURRENT_PROJECT_NOT_SAVED")):
            with self.subTest(expected=expected), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=results)), \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch(ARGS)
            self.assertEqual(result["error_code"], expected)
            post.assert_not_awaited()

    async def test_save_failure_reports_opened_template(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "exists": True}, {"success": True, "exists": False},
                {"success": True, "isSaved": True}, CAPABILITIES,
                {"success": True, "isSaved": True, "projectPath": "C:/projects/template.ymmp"}])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True}, {"success": False, "error_code": "DIRECTORY_NOT_FOUND"}])):
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "TEMPLATE_SAVE_FAILED")
        self.assertEqual(result["opened_template_path"], ARGS["template_path"])

    async def test_open_without_verified_template_never_saves_previous_project(self):
        states = iter([{"success": True, "exists": True}, {"success": True, "exists": False},
                       {"success": True, "isSaved": True}, CAPABILITIES])
        async def get(_path):
            return next(states, {"success": True, "isSaved": True,
                                 "projectPath": "C:/projects/old.ymmp"})
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post, \
             patch.object(server.asyncio, "sleep", new_callable=AsyncMock):
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "TEMPLATE_OPEN_NOT_VERIFIED")
        self.assertEqual(post.await_count, 1)

    async def test_plugin_without_guard_does_not_open_template(self):
        for capability in ({"success": True, "features": {"project_open_save_as": True}},
                           {"success": True, "features": {"guarded_template_save": 1}},
                           {**CAPABILITIES, "error_code": "UNAVAILABLE"}):
            with self.subTest(capability=capability), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                     {"success": True, "exists": True}, {"success": True, "exists": False},
                     {"success": True, "isSaved": True}, capability])) as get, \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch(ARGS)
            self.assertEqual(result["error_code"], "TEMPLATE_GUARDED_SAVE_UNAVAILABLE")
            self.assertEqual(get.await_args.args, ("/capabilities",))
            post.assert_not_awaited()

    async def test_template_switch_at_save_does_not_verify_or_retry(self):
        gets = [{"success": True, "exists": True}, {"success": True, "exists": False},
                {"success": True, "isSaved": True}, CAPABILITIES,
                {"success": True, "isSaved": True, "projectPath": ARGS["template_path"]}]
        conflict = {"success": False, "error_code": "TEMPLATE_SOURCE_CHANGED"}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)) as get, \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True}, conflict])) as post:
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "TEMPLATE_SOURCE_CHANGED")
        self.assertEqual(result["template_path"], ARGS["template_path"])
        self.assertNotIn("opened_template_path", result)
        self.assertEqual(get.await_count, len(gets))
        self.assertEqual(post.await_count, 2)
        self.assertEqual(post.await_args.args, ("/project/save-template-as", {
            "path": ARGS["path"], "expected_template_path": ARGS["template_path"]}))

    async def test_same_path_rejected_before_host_call(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({**ARGS, "path": "c:\\projects\\TEMPLATE.ymmp"})
        get.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
