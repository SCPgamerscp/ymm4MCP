"""MCP integration regression for the read-only ripple planner (#61)."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class RippleDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_dispatch_only_reads_live_items(self):
        snapshot = {"success": True, "items": [
            {"item_id": "native:cut", "frame": 10, "length": 10,
             "layer": 1, "revision": "r1"},
            {"item_id": "native:after", "frame": 20, "length": 10,
             "layer": 1, "revision": "r2"}]}
        with (patch.object(server, "ymm4_get", new_callable=AsyncMock, return_value=snapshot) as get,
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "plan_ripple_delete", "start_frame": 10, "end_frame": 20,
                "layers": [1],
                "expected_revisions": {"native:cut": "r1", "native:after": "r2"}})
        self.assertTrue(result["success"])
        self.assertFalse(result["executable"])
        self.assertEqual(result["delete_count"], 1)
        self.assertEqual(result["move_count"], 1)
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_invalid_range_fails_before_host_read(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "plan_ripple_delete",
                                       "start_frame": 20, "end_frame": 10})
            get.assert_not_awaited()

    async def test_stale_revision_fails_without_post(self):
        snapshot = {"items": [{"item_id": "native:a", "frame": 30,
                               "length": 10, "layer": 1, "revision": "new"}]}
        with (patch.object(server, "ymm4_get", new_callable=AsyncMock, return_value=snapshot),
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "plan_ripple_delete", "start_frame": 10, "end_frame": 20,
                "expected_revisions": {"native:a": "old"}})
        self.assertEqual(result["error_code"], "RIPPLE_CONFLICT")
        post.assert_not_awaited()


    async def test_verify_ripple_delete_reads_only_current_snapshot(self):
        before = {"success": True, "count": 2, "items": [
            {"item_id": "cut", "frame": 10, "length": 10, "layer": 1, "revision": "r1"},
            {"item_id": "after", "frame": 20, "length": 10, "layer": 1, "revision": "r2"}]}
        after = {"success": True, "count": 1, "items": [
            {"item_id": "after", "frame": 10, "length": 10, "layer": 1, "revision": "r3"}]}
        with (patch.object(server, "ymm4_get", new_callable=AsyncMock, return_value=after) as get,
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "verify_ripple_delete", "before_snapshot": before,
                "start_frame": 10, "end_frame": 20, "layers": [1],
                "expected_revisions": {"cut": "r1", "after": "r2"}})
        self.assertTrue(result["passed"])
        self.assertFalse(result["atomicity_verified"])
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_verify_ripple_delete_detects_missing_shift(self):
        before = {"count": 1, "items": [
            {"item_id": "after", "frame": 20, "length": 10, "layer": 1, "revision": "r1"}]}
        with patch.object(server, "ymm4_get", new_callable=AsyncMock, return_value=before):
            result = await server.dispatch({
                "action": "verify_ripple_delete", "before_snapshot": before,
                "start_frame": 10, "end_frame": 20,
                "expected_revisions": {"after": "r1"}})
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["code"], "RIPPLE_FIELD_MISMATCH")

    async def test_verifying_removal_of_last_item_requires_a_readable_inventory(self):
        before = {"count": 1, "items": [
            {"item_id": "cut", "frame": 10, "length": 10, "layer": 1, "revision": "r1"}]}
        args = {"action": "verify_ripple_delete", "before_snapshot": before,
                "start_frame": 10, "end_frame": 20, "expected_revisions": {"cut": "r1"}}
        outcomes = ({"success": False, "error_code": "ITEMS_UNAVAILABLE"},
                    {"count": 0, "items": []})
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=outcomes)):
            unavailable = await server.dispatch(args)
            empty = await server.dispatch(args)
        self.assertFalse(unavailable["passed"])
        self.assertEqual(unavailable["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")
        self.assertTrue(empty["passed"])

    async def test_verify_ripple_delete_requires_before_snapshot(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "verify_ripple_delete",
                                       "start_frame": 10, "end_frame": 20})
            get.assert_not_awaited()



if __name__ == "__main__":
    unittest.main()
