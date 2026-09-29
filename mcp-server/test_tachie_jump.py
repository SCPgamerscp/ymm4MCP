"""A tachie jump should plan safely and verify all applied keyframes."""
import unittest
from unittest.mock import AsyncMock, patch

import server

ITEMS = {"success": True, "items": [{"item_id": "tachie:1", "type": "TachieItem",
                                      "revision": "r1", "length": 100}]}
Y = {"success": True, "revision": "r1", "animations": [
    {"prop": "Y", "keyframes": [{"at": 0, "value": 100}]}]}


class JumpTests(unittest.IsolatedAsyncioTestCase):
    async def test_reaction_can_start_at_voice_frame_relative_to_tachie(self):
        items = {"success": True, "items": [
            {**ITEMS["items"][0], "frame": 100},
            {"item_id": "voice:1", "type": "VoiceItem", "frame": 120, "length": 30}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[items, Y])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "jump_tachie", "item_id": "tachie:1",
                                            "voice_item_id": "voice:1", "at": 3})
        self.assertEqual(result["at"], 23)
        self.assertEqual(result["keyframes"][0]["at"], 23)
        post.assert_not_awaited()

    async def test_missing_voice_stops_before_keyframe_lookup(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "shake_tachie", "item_id": "tachie:1",
                                            "voice_item_id": "missing", "dry_run": False})
        self.assertEqual(result["error_code"], "VOICE_ITEM_NOT_FOUND")
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_shake_uses_x_axis_and_bounded_five_point_pattern(self):
        x = {"success": True, "revision": "r1", "animations": [
            {"prop": "X", "keyframes": [{"at": 0, "value": 100}]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, x])) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "shake_tachie", "item_id": "tachie:1",
                                            "at": 20, "duration_frames": 8, "shake_distance": 25})
        self.assertEqual(result["keyframes"], [
            {"at": 20, "value": 100}, {"at": 22, "value": 125},
            {"at": 24, "value": 75}, {"at": 26, "value": 125},
            {"at": 28, "value": 100}])
        self.assertIn("prop=X", get.await_args_list[1].args[0])
        post.assert_not_awaited()

    async def test_shake_rejects_existing_x_animation(self):
        x = {"success": True, "revision": "r1", "animations": [
            {"prop": "X", "keyframes": [{"at": 0, "value": 100}, {"at": 3, "value": 120}]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, x])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "shake_tachie", "item_id": "tachie:1",
                                            "dry_run": False})
        self.assertEqual(result["error_code"], "TACHIE_HAS_EXISTING_X_KEYFRAMES")
        post.assert_not_awaited()

    async def test_shake_applies_five_x_keys_with_revision_chain(self):
        x = {"success": True, "revision": "r1", "animations": [
            {"prop": "X", "keyframes": [{"at": 0, "value": 100}]}]}
        points = [{"at": at, "value": value} for at, value in
                  [(20, 100), (22, 125), (24, 75), (26, 125), (28, 100)]]
        checked = {"success": True, "revision": "r6", "animations": [
            {"prop": "X", "keyframes": [{"at": 0, "value": 100}, *points]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                ITEMS, x, {"success": True, "isSaved": True}, checked])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True, "backup_path": "C:/backup.ymmp"},
                 *({"success": True, "revision": f"r{i}"} for i in range(2, 7))])) as post:
            result = await server.dispatch({"action": "shake_tachie", "item_id": "tachie:1",
                                            "at": 20, "duration_frames": 8,
                                            "shake_distance": 25, "dry_run": False})
        self.assertTrue(result["verified"])
        self.assertEqual([call.args[1]["expected_revision"] for call in post.await_args_list[1:]],
                         ["r1", "r2", "r3", "r4", "r5"])
        self.assertTrue(all(call.args[1]["prop"] == "X" for call in post.await_args_list[1:]))

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
