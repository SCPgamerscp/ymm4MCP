"""BGM ducking must reject contradictory host responses before trusting edits."""
import unittest
from unittest.mock import AsyncMock, patch

import server


ITEMS = {"items": [
    {"item_id": "bgm:1", "revision": "r1", "type": "AudioItem", "frame": 0, "length": 100},
    {"item_id": "voice:1", "type": "VoiceItem", "frame": 20, "length": 10},
]}
KEYS = {"success": True, "revision": "r1", "animations": [
    {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}]}]}
PROJECT = {"success": True, "isSaved": True}
CHECKPOINT = {"success": True, "backup_path": "C:/backup/bgm.ymmp"}
ARGS = {"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False}


class BgmHostResponseTests(unittest.IsolatedAsyncioTestCase):
    async def test_snapshot_error_markers_and_invalid_success_reject_without_writes(self):
        for snapshot in (
            {**ITEMS, "success": True, "error_code": None},
            {**ITEMS, "success": True, "error": "unavailable"},
            {**ITEMS, "success": "true"},
            None,
        ):
            with self.subTest(snapshot=snapshot):
                with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)) as get, \
                     patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                    result = await server.dispatch(ARGS)
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                get.assert_awaited_once_with("/items")
                post.assert_not_awaited()

    async def test_keyframe_read_error_marker_rejects_before_backup(self):
        for marker in ("error", "error_code"):
            with self.subTest(marker=marker):
                with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                        ITEMS, {**KEYS, marker: None}])), \
                     patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                    result = await server.dispatch(ARGS)
                self.assertEqual(result["error_code"], "BGM_KEYFRAMES_UNAVAILABLE")
                post.assert_not_awaited()

    async def test_project_error_marker_rejects_before_backup(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                ITEMS, KEYS, {**PROJECT, "error_code": "PROJECT_UNAVAILABLE"}])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "PROJECT_SAVE_REQUIRED")
        post.assert_not_awaited()

    async def test_checkpoint_error_marker_rejects_before_keyframe_write(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, KEYS, PROJECT])), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={
                 **CHECKPOINT, "error_code": None})) as post:
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "BGM_BACKUP_REQUIRED")
        self.assertEqual(post.await_count, 1)
        self.assertEqual(post.await_args.args[0], "/edits/checkpoint")

    async def test_mutation_error_marker_is_reported_as_unknown_outcome(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                ITEMS, KEYS, PROJECT, ITEMS])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 CHECKPOINT, {"success": True, "revision": "r2",
                              "error_code": None}])) as post:
            result = await server.dispatch(ARGS)
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        self.assertEqual(result["applied"], [])
        self.assertEqual(result["backup_path"], CHECKPOINT["backup_path"])
        self.assertEqual(post.await_count, 2)


if __name__ == "__main__":
    unittest.main()
