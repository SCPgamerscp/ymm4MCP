"""MCP dispatch regression: capture and verify against server-owned evidence."""
import unittest
from unittest.mock import AsyncMock, patch

import server


def snapshot(*items):
    return {"success": True, "count": len(items), "items": list(items)}


BEFORE = snapshot(
    {"item_id": "cut", "frame": 10, "length": 10, "layer": 1, "revision": "r1"},
    {"item_id": "later", "frame": 20, "length": 10, "layer": 1, "revision": "r2"})
AFTER = snapshot(
    {"item_id": "later", "frame": 10, "length": 10, "layer": 1, "revision": "r3"})


class RippleEvidenceDispatchTests(unittest.IsolatedAsyncioTestCase):
    def plan_args(self):
        return {"action": "plan_ripple_delete", "start_frame": 10,
                "end_frame": 20, "layers": [1],
                "expected_revisions": {"cut": "r1", "later": "r2"},
                "capture_evidence": True}

    async def test_server_captured_preedit_inventory_is_used(self):
        with (patch.object(server, "ymm4_get", new_callable=AsyncMock,
                           side_effect=[BEFORE, AFTER]) as get,
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            plan = await server.dispatch(self.plan_args())
            self.assertTrue(plan["success"])
            self.assertFalse(plan["executable"])
            self.assertEqual(plan["evidence_source"], "server_captured_host_inventory")
            result = await server.dispatch(
                {"action": "verify_ripple_delete", "plan_id": plan["plan_id"]})
        self.assertTrue(result["passed"])
        self.assertFalse(result["atomicity_verified"])
        self.assertEqual(result["before_sha256"], plan["before_sha256"])
        self.assertEqual(get.await_count, 2)
        post.assert_not_awaited()

    async def test_plan_id_cannot_be_combined_with_forged_before(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "verify_ripple_delete",
                                       "plan_id": "arbitrary",
                                       "before_snapshot": AFTER})
            get.assert_not_awaited()

    async def test_unknown_plan_id_does_not_read_host(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            result = await server.dispatch(
                {"action": "verify_ripple_delete", "plan_id": "not-a-real-id"})
            self.assertEqual(result["error_code"], "RIPPLE_EVIDENCE_UNAVAILABLE")
            get.assert_not_awaited()

    async def test_incomplete_host_snapshot_cannot_be_captured(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock,
                          return_value={"items": BEFORE["items"]}):
            result = await server.dispatch(self.plan_args())
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "RIPPLE_EVIDENCE_UNAVAILABLE")
        self.assertNotIn("plan_id", result)

    async def test_invalid_capture_flag_rejected_before_host_read(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({**self.plan_args(), "capture_evidence": "yes"})
            get.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
