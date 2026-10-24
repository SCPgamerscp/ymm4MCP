"""A host error response must never authorize an optimistic script write."""
import unittest
from unittest.mock import AsyncMock

import editplan
import script_guard


class ScriptSnapshotHostErrorTests(unittest.IsolatedAsyncioTestCase):
    def args(self):
        return {"lines": [{"text": "hello", "character": "A"}],
                "expected_snapshot_hash": editplan.snapshot_hash([])}

    async def test_error_with_valid_items_does_not_write(self):
        for snapshot in (
            {"success": True, "items": [], "error": "host unavailable"},
            {"items": [], "error": "host unavailable"},
        ):
            with self.subTest(snapshot=snapshot):
                get = AsyncMock(return_value=snapshot)
                add_script = AsyncMock(return_value={"success": True, "added": 1})
                result = await script_guard.apply(self.args(), get, add_script)
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                self.assertEqual(result["details"], snapshot)
                get.assert_awaited_once_with("/items")
                add_script.assert_not_awaited()

    async def test_error_code_and_nonboolean_success_never_authorize_write(self):
        for snapshot in (
            {"success": True, "items": [], "error_code": "HOST_UNAVAILABLE"},
            {"success": True, "items": [], "error_code": None},
            {"success": "true", "items": []},
            {"success": 1, "items": []},
            {"success": None, "items": []},
        ):
            with self.subTest(snapshot=snapshot):
                add_script = AsyncMock()
                result = await script_guard.apply(
                    self.args(), AsyncMock(return_value=snapshot), add_script)
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                self.assertEqual(result["details"], snapshot)
                add_script.assert_not_awaited()

    async def test_clean_matching_snapshot_still_allows_write(self):
        args = self.args()
        get = AsyncMock(return_value={"success": True, "items": []})
        add_script = AsyncMock(return_value={"success": True, "added": 1})
        result = await script_guard.apply(args, get, add_script)
        self.assertEqual(result, {"success": True, "added": 1})
        add_script.assert_awaited_once_with(args, expected_snapshot=args["expected_snapshot_hash"])


if __name__ == "__main__":
    unittest.main()
