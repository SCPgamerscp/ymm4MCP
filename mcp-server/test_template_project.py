"""Template creation must guard source and destination before opening YMM4 projects."""
import unittest
from unittest.mock import AsyncMock, patch

import server


ARGS = {"action": "create_from_template", "template_path": "C:/projects/template.ymmp",
        "path": "C:/projects/new.ymmp"}


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
                {"success": True, "isSaved": True},
                {"success": True, "isSaved": True, "projectPath": "c:\\projects\\new.ymmp"}])), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            result = await server.dispatch(ARGS)
        self.assertTrue(result["verified"])
        self.assertEqual([call.args[0] for call in post.await_args_list],
                         ["/project/open", "/project/save-as"])
        self.assertEqual(post.await_args_list[-1].args[1]["overwrite"], False)

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
                {"success": True, "isSaved": True}])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True}, {"success": False, "error_code": "DIRECTORY_NOT_FOUND"}])):
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "TEMPLATE_SAVE_FAILED")
        self.assertEqual(result["opened_template_path"], ARGS["template_path"])

    async def test_same_path_rejected_before_host_call(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({**ARGS, "path": "c:\\projects\\TEMPLATE.ymmp"})
        get.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
