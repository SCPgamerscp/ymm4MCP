"""MCP backup restoration must never fall through to item rollback."""

import unittest
from unittest.mock import AsyncMock, patch

import httpx

import server


CHECKPOINT = "cp_12345678"
ROLLBACK = {"action": "control", "sub_action": "rollback", "checkpoint_id": CHECKPOINT}
RESTORE = {**ROLLBACK, "restore_backup": True, "confirm_restore_backup": True}
RESTORED = {
    "success": True,
    "action": "restore_backup",
    "checkpoint_id": CHECKPOINT,
    "backup_path": "C:/checkpoints/cp_12345678.ymmp",
    "backup_sha256": "A" * 64,
    "original_project_path": "C:/projects/original.ymmp",
}
OPENED = {"success": True, "projectPath": "C:/CHECKPOINTS/cp_12345678.ymmp", "isSaved": True}


class CheckpointRestoreDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_confirmed_restore_uses_only_the_restore_endpoint(self):
        with (
            patch.object(server, "ymm4_post", AsyncMock(return_value=RESTORED)) as post,
            patch.object(server, "ymm4_get", AsyncMock(return_value=OPENED)) as get,
        ):
            self.assertEqual(await server.dispatch(RESTORE), RESTORED)
        post.assert_awaited_once_with("/edits/restore-backup", {
            "checkpoint_id": CHECKPOINT, "confirm_restore_backup": True})
        get.assert_awaited_once_with("/project")

    async def test_success_receipt_does_not_certify_a_different_or_unverified_open_project(self):
        for state in (
            {**OPENED, "projectPath": RESTORED["original_project_path"]},
            {**OPENED, "isSaved": False},
            {**OPENED, "isSaved": 1},
            {**OPENED, "error_code": "UNAVAILABLE"},
            None,
            httpx.ReadTimeout("project readback timed out"),
            ValueError("invalid project JSON"),
        ):
            with self.subTest(state=state):
                with (
                    patch.object(server, "ymm4_post", AsyncMock(return_value=RESTORED)) as post,
                    patch.object(server, "ymm4_get", AsyncMock(side_effect=[state])) as get,
                ):
                    result = await server.dispatch(RESTORE)
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "RESTORE_OUTCOME_UNKNOWN")
                self.assertIs(result["outcome_unknown"], True)
                self.assertEqual(result["backup_path"], RESTORED["backup_path"])
                post.assert_awaited_once_with("/edits/restore-backup", {
                    "checkpoint_id": CHECKPOINT, "confirm_restore_backup": True})
                get.assert_awaited_once_with("/project")

    async def test_unconfirmed_or_invalid_restore_sends_no_edit(self):
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            unconfirmed = await server.dispatch({**ROLLBACK, "restore_backup": True})
            self.assertEqual(unconfirmed["error_code"], "RESTORE_CONFIRMATION_REQUIRED")
            for invalid in ({"confirm_restore_backup": True},
                            {"restore_backup": "true", "confirm_restore_backup": True},
                            {"restore_backup": True, "confirm_restore_backup": 1}):
                with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                    await server.dispatch({**ROLLBACK, **invalid})
        post.assert_not_awaited()

    async def test_old_plugin_does_not_fall_back_to_item_deletion(self):
        unsupported = {"success": False, "error_code": "NOT_FOUND"}
        with patch.object(server, "ymm4_post", AsyncMock(return_value=unsupported)) as post:
            result = await server.dispatch(RESTORE)
        self.assertEqual(result["error_code"], "CHECKPOINT_RESTORE_UNAVAILABLE")
        post.assert_awaited_once_with("/edits/restore-backup", {
            "checkpoint_id": CHECKPOINT, "confirm_restore_backup": True})

    async def test_structured_restore_failure_is_preserved(self):
        for code in ("UNSAVED_CHANGES", "CHECKPOINT_BACKUP_UNAVAILABLE", "RESTORE_OUTCOME_UNKNOWN"):
            rejected = {"success": False, "error_code": code,
                        "outcome_unknown": code == "RESTORE_OUTCOME_UNKNOWN"}
            with self.subTest(code=code), \
                 patch.object(server, "ymm4_post", AsyncMock(return_value=rejected)):
                self.assertEqual(await server.dispatch(RESTORE), rejected)

    async def test_partial_or_mismatched_receipt_does_not_report_success(self):
        for reply in (None, {"success": True, "removed": 3},
                      {**RESTORED, "checkpoint_id": "cp_00000000"},
                      {**RESTORED, "backup_sha256": "abc"},
                      {**RESTORED, "backup_path": "C:/PROJECTS/original.ymmp"}):
            with self.subTest(reply=reply), \
                 patch.object(server, "ymm4_post", AsyncMock(return_value=reply)) as post:
                result = await server.dispatch(RESTORE)
            self.assertEqual(result["error_code"], "RESTORE_OUTCOME_UNKNOWN")
            self.assertTrue(result["outcome_unknown"])
            post.assert_awaited_once()

    async def test_only_explicit_false_outcome_unknown_can_confirm_restore(self):
        for flag in (True, None, 0, "false", False):
            receipt = {**RESTORED, "outcome_unknown": flag}
            with (
                self.subTest(flag=flag),
                patch.object(server, "ymm4_post", AsyncMock(return_value=receipt)) as post,
                patch.object(server, "ymm4_get", AsyncMock(return_value=OPENED)) as get,
            ):
                result = await server.dispatch(RESTORE)
                if flag is False:
                    self.assertEqual(result, receipt)
                    get.assert_awaited_once_with("/project")
                else:
                    self.assertIs(result["success"], False)
                    self.assertEqual(result["error_code"], "RESTORE_OUTCOME_UNKNOWN")
                    get.assert_not_awaited()
            post.assert_awaited_once()

    async def test_restore_timeout_is_indeterminate_and_is_not_retried(self):
        with patch.object(server, "ymm4_post",
                          AsyncMock(side_effect=httpx.ReadTimeout("editing host timed out"))) as post:
            result = await server.dispatch(RESTORE)
        self.assertEqual(result["error_code"], "RESTORE_OUTCOME_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        post.assert_awaited_once()
