"""Host status must be explicit when supplied, so QA cannot pass a failed readback."""
import unittest
from unittest.mock import AsyncMock

from quality_guard import verify_snapshot


ITEMS = [{"item_id": "voice:1", "frame": 0, "length": 20, "layer": 1}]


class SnapshotHostStatusTests(unittest.IsolatedAsyncioTestCase):
    async def test_non_boolean_or_false_success_is_unavailable(self):
        for status in (False, None, 0, 1, "false", "true", [], {}):
            with self.subTest(status=status):
                get = AsyncMock(return_value={"success": status, "items": ITEMS})
                result = await verify_snapshot(get, ITEMS)
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                get.assert_awaited_once_with("/items")

    async def test_explicit_success_and_legacy_omitted_status_still_pass(self):
        for report in ({"success": True, "items": ITEMS}, {"items": ITEMS}):
            with self.subTest(report=report):
                self.assertIsNone(await verify_snapshot(AsyncMock(return_value=report), ITEMS))

    async def test_host_error_never_passes_even_with_success_true(self):
        result = await verify_snapshot(AsyncMock(return_value={
            "success": True, "items": ITEMS, "error": "read failed"}), ITEMS)
        self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")

    async def test_host_error_code_marker_never_authorizes_qa(self):
        for code in ("ITEMS_UNAVAILABLE", None):
            with self.subTest(code=code):
                response = {"success": True, "items": ITEMS, "error_code": code}
                result = await verify_snapshot(AsyncMock(return_value=response), ITEMS)
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                self.assertEqual(result["details"], response)

    async def test_valid_status_does_not_hide_snapshot_conflict(self):
        different = [{**ITEMS[0], "frame": 1}]
        result = await verify_snapshot(AsyncMock(return_value={
            "success": True, "items": different}), ITEMS)
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")

    async def test_host_exception_stays_unknown(self):
        result = await verify_snapshot(AsyncMock(side_effect=TimeoutError("timeout")), ITEMS)
        self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")


if __name__ == "__main__":
    unittest.main()
