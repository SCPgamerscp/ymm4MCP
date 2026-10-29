"""Guarded script execution must not certify malformed host results."""
import unittest
from unittest.mock import AsyncMock

import editplan
import script_guard


class ScriptResultIntegrityTests(unittest.IsolatedAsyncioTestCase):
    def args(self, *, dry_run=False):
        return {"lines": [{"text": "hello", "character": "A"}],
                "expected_snapshot_hash": editplan.snapshot_hash([]),
                "dry_run": dry_run}

    async def run_script(self, response, *, dry_run=False):
        args = self.args(dry_run=dry_run)
        add_script = AsyncMock(return_value=response)
        result = await script_guard.apply(
            args, AsyncMock(return_value={"success": True, "items": []}), add_script)
        add_script.assert_awaited_once_with(
            args, expected_snapshot=args["expected_snapshot_hash"])
        return result

    async def test_non_object_host_results_fail_closed(self):
        for response in (None, True, "ok", [], 0):
            with self.subTest(response=response):
                result = await self.run_script(response)
                self.assertEqual(result["error_code"], "SCRIPT_RESULT_INVALID")
                self.assertFalse(result["success"])

    async def test_ambiguous_success_fields_fail_closed(self):
        for response in ({"added": 1}, {"success": 1}, {"success": "true"},
                         {"success": None}, {"success": True, "error": "failure"},
                         {"success": True, "error_code": None}):
            with self.subTest(response=response):
                result = await self.run_script(response)
                self.assertEqual(result["error_code"], "SCRIPT_RESULT_INVALID")

    async def test_failed_dry_run_is_not_certified_by_snapshot_hash(self):
        failed = {"success": False, "error_code": "CHARACTER_NOT_FOUND"}
        result = await self.run_script(failed, dry_run=True)
        self.assertEqual(result, failed)
        self.assertNotIn("snapshot_hash", result)

    async def test_successful_dry_run_still_gets_snapshot_hash(self):
        result = await self.run_script({"success": True, "dry_run": True}, dry_run=True)
        self.assertEqual(result["snapshot_hash"], editplan.snapshot_hash([]))

    async def test_successful_apply_result_is_unchanged(self):
        good = {"success": True, "added": 1}
        self.assertEqual(await self.run_script(good), good)


if __name__ == "__main__":
    unittest.main()
