"""Malformed reference and host snapshots must fail closed without host writes."""
import unittest
from unittest.mock import AsyncMock

from quality_guard import verify_snapshot


VALID = [{"item_id": "v", "frame": 0, "length": 30, "layer": 1}]


class InvalidSnapshotTests(unittest.IsolatedAsyncioTestCase):
    async def test_invalid_reference_snapshot_is_reported_without_host_read(self):
        for invalid in (None, {}, [None], [{"item_id": object()}]):
            with self.subTest(invalid=repr(invalid)):
                get = AsyncMock()
                result = await verify_snapshot(get, invalid)
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                get.assert_not_awaited()

    async def test_invalid_host_items_cannot_match_reference(self):
        for invalid in ([None], [{"item_id": object()}]):
            with self.subTest(invalid=repr(invalid)):
                get = AsyncMock(return_value={"success": True, "items": invalid})
                result = await verify_snapshot(get, VALID)
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                get.assert_awaited_once_with("/items")

    async def test_valid_reference_still_detects_conflicts(self):
        changed = [{**VALID[0], "frame": 1}]
        result = await verify_snapshot(AsyncMock(return_value={"items": changed}), VALID)
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")


if __name__ == "__main__":
    unittest.main()
