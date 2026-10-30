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


if __name__ == "__main__":
    unittest.main()
