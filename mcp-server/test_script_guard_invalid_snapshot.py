"""A failed optimistic script snapshot read must never start editing."""
import unittest
from unittest.mock import AsyncMock

import editplan
import script_guard


ARGS = {"lines": [{"text": "hello", "character": "A"}],
        "expected_snapshot_hash": editplan.snapshot_hash([])}


class ScriptGuardInvalidSnapshotTests(unittest.IsolatedAsyncioTestCase):
    async def test_host_read_exception_prevents_writes(self):
        add_script = AsyncMock()
        result = await script_guard.apply(ARGS, AsyncMock(side_effect=TimeoutError("timeout")), add_script)
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
        add_script.assert_not_awaited()

    async def test_invalid_host_items_prevent_writes(self):
        for items in ([None], [{"frame": object()}]):
            with self.subTest(items=repr(items)):
                add_script = AsyncMock()
                result = await script_guard.apply(ARGS, AsyncMock(return_value={
                    "success": True, "items": items}), add_script)
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                add_script.assert_not_awaited()

    async def test_host_status_error_still_returns_items_unavailable(self):
        add_script = AsyncMock()
        result = await script_guard.apply(ARGS, AsyncMock(return_value={
            "success": False, "items": []}), add_script)
        self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
        add_script.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
