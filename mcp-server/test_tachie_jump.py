"""A tachie jump should plan safely and verify all applied keyframes."""
import unittest
from unittest.mock import AsyncMock, patch

import server

ITEMS = {"success": True, "items": [{"item_id": "tachie:1", "type": "TachieItem",
                                      "revision": "r1", "length": 100}]}
Y = {"success": True, "revision": "r1", "animations": [
    {"prop": "Y", "keyframes": [{"at": 0, "value": 100}]}]}


class JumpTests(unittest.IsolatedAsyncioTestCase):
    async def test_dry_run_plans_without_mutation(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, Y])) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "jump_tachie", "item_id": "tachie:1",
                                            "at": 20, "duration_frames": 12, "jump_height": 40})
        self.assertEqual(result["keyframes"], [{"at": 20, "value": 100},
                                               {"at": 26, "value": 60},
                                               {"at": 32, "value": 100}])
        self.assertEqual(get.await_count, 2)
        post.assert_not_awaited()

    async def test_existing_y_animation_stops_before_mutation(self):
        existing = {**Y, "animations": [{"prop": "Y", "keyframes": [
            {"at": 0, "value": 100}, {"at": 5, "value": 50}]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, existing])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "jump_tachie", "item_id": "tachie:1",
                                            "dry_run": False})
        self.assertEqual(result["error_code"], "TACHIE_HAS_EXISTING_Y_KEYFRAMES")
        post.assert_not_awaited()

    async def test_applies_with_backup_revision_chain_and_verification(self):
        points = [{"at": 20, "value": 100}, {"at": 26, "value": 60}, {"at": 32, "value": 100}]
        checked = {"success": True, "revision": "r4", "animations": [
            {"prop": "Y", "keyframes": [{"at": 0, "value": 100}, *points]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, Y, {
                "success": True, "isSaved": True}, checked])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True, "backup_path": "C:/backup.ymmp"},
                 {"success": True, "revision": "r2"},
                 {"success": True, "revision": "r3"},
                 {"success": True, "revision": "r4"}])) as post:
            result = await server.dispatch({"action": "jump_tachie", "item_id": "tachie:1",
                                            "at": 20, "dry_run": False})
        self.assertTrue(result["verified"])
        self.assertEqual([call.args[1]["expected_revision"] for call in post.await_args_list[1:]],
                         ["r1", "r2", "r3"])


if __name__ == "__main__":
    unittest.main()
