"""Voice-aware BGM keyframe planning and safe dispatch."""
import unittest
from unittest.mock import AsyncMock, patch

import ducking
import server


ITEMS = {"items": [
    {"item_id": "bgm:1", "revision": "r1", "type": "AudioItem", "frame": 0, "length": 100},
    {"item_id": "voice:1", "type": "VoiceItem", "frame": 20, "length": 10},
]}
KEYS = {"success": True, "revision": "r1", "animations": [
    {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}]}]}


class DuckingPlanTests(unittest.TestCase):
    def test_voice_span_creates_attack_hold_and_release(self):
        planned = ducking.plan(ITEMS["items"][0], ITEMS["items"][1:], 100)
        self.assertEqual(planned, [{"at": 15, "value": 100.0}, {"at": 20, "value": 30.0},
                                   {"at": 30, "value": 30.0}, {"at": 40, "value": 100.0}])

    def test_close_spans_merge_and_clipped_end_does_not_restore_early(self):
        bgm = {"frame": 100, "length": 30}
        voice = [{"frame": 95, "length": 10}, {"frame": 110, "length": 30}]
        planned = ducking.plan(bgm, voice, 80, attack_frames=5, release_frames=10)
        self.assertEqual(planned, [{"at": 0, "value": 24.0}])

    def test_short_tail_stays_ducked_through_bgm_end(self):
        bgm = {"frame": 0, "length": 100}
        voices = [{"frame": 85, "length": 10}]
        self.assertEqual(ducking.plan(bgm, voices, 100, release_frames=10),
                         [{"at": 80, "value": 100.0}, {"at": 85, "value": 30.0},
                          {"at": 95, "value": 30.0}])
        self.assertEqual(ducking.plan(bgm, [{"frame": 75, "length": 10}], 100,
                                      release_frames=10)[-1], {"at": 95, "value": 100.0})

    def test_rejects_invalid_ratios_and_items(self):
        for ratio in (0, 1, True, float("nan")):
            with self.subTest(ratio=ratio), self.assertRaises(ValueError):
                ducking.plan(ITEMS["items"][0], ITEMS["items"][1:], 100, ratio=ratio)


class DuckingDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_selected_voice_outside_bgm_rejects_before_keyframe_access(self):
        distant = {"items": [ITEMS["items"][0],
                             {**ITEMS["items"][1], "frame": 100}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=distant)) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "voice_item_ids": ["voice:1"], "dry_run": False})
        self.assertEqual(result["error_code"], "VOICE_OUTSIDE_BGM")
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_voice_selection_is_fail_closed(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "voice_item_ids": ["missing"]})
        self.assertEqual(result["error_code"], "VOICE_ITEM_NOT_FOUND")
        get.assert_awaited_once()
        post.assert_not_awaited()
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, KEYS])):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "voice_item_ids": ["voice:1"]})
        self.assertEqual(result["voice_count"], 1)

    async def test_dry_run_reads_only_and_rejects_existing_keys(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, KEYS])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
        self.assertTrue(result["dry_run"])
        self.assertEqual(len(result["keyframes"]), 4)
        post.assert_not_awaited()
        custom = {**KEYS, "animations": [{"prop": "Volume", "keyframes": [
            {"at": 0, "value": 100}, {"at": 5, "value": 50}]}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, custom])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_HAS_EXISTING_KEYFRAMES")
        post.assert_not_awaited()

    async def test_apply_uses_revision_chain_and_verifies(self):
        planned = ducking.plan(ITEMS["items"][0], ITEMS["items"][1:], 100)
        verified = {"success": True, "animations": [{"prop": "Volume", "keyframes": planned}]}
        gets = [ITEMS, KEYS, {"success": True, "isSaved": True}, verified]
        posts = [{"success": True, "backup_path": "C:/backup/cp.ymmp"}] + [
            {"success": True, "revision": f"r{i}"} for i in range(2, 6)]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=posts)) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertTrue(result["verified"])
        self.assertEqual([call.args[1]["expected_revision"] for call in post.await_args_list[1:]],
                         ["r1", "r2", "r3", "r4"])

    async def test_partial_failure_exposes_backup_and_applied_points(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                 ITEMS, KEYS, {"success": True, "isSaved": True}])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True, "backup_path": "C:/backup/cp.ymmp"},
                 {"success": True, "revision": "r2"},
                 {"success": False, "error_code": "KEYFRAME_METHOD_UNAVAILABLE"}])):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_DUCKING_PARTIAL")
        self.assertEqual(len(result["applied"]), 1)
        self.assertEqual(result["backup_path"], "C:/backup/cp.ymmp")

    async def test_verification_failure_preserves_backup_path(self):
        posts = [{"success": True, "backup_path": "C:/backup/cp.ymmp"}] + [
            {"success": True, "revision": f"r{i}"} for i in range(2, 6)]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                 ITEMS, KEYS, {"success": True, "isSaved": True}, TimeoutError("read timed out")])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=posts)):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_DUCKING_VERIFY_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        self.assertEqual(result["backup_path"], "C:/backup/cp.ymmp")


if __name__ == "__main__":
    unittest.main()
