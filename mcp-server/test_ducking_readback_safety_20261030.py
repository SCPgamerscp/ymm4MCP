"""Fail-closed readback and host-state regression tests for BGM ducking."""
import unittest
from unittest.mock import AsyncMock, patch

import ducking
import server


BGM = {"item_id": "bgm:1", "revision": "r1", "type": "AudioItem",
       "frame": 0, "length": 100}
VOICE = {"item_id": "voice:1", "type": "VoiceItem", "frame": 20, "length": 10}
ITEMS = {"items": [BGM, VOICE]}
KEYS = {"success": True, "revision": "r1", "animations": [
    {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}]}]}


def readback(value=30.0, *, revision="r2"):
    return {"success": True, "revision": revision, "animations": [
        {"prop": "Volume", "keyframes": [
            {"at": 0, "value": 100.0}, {"at": 20, "value": value}]}]}


class ReadbackIntegrityTests(unittest.TestCase):
    def test_exact_curve_still_verifies(self):
        self.assertTrue(ducking.verify_keyframes(
            readback(), "r2", [{"at": 20, "value": 30.0}], 100.0))

    def test_small_float_roundoff_is_accepted(self):
        self.assertTrue(ducking.verify_keyframes(
            readback(30.000000001), "r2", [{"at": 20, "value": 30.0}], 100.0))

    def test_zero_host_volume_cannot_masquerade_as_tiny_duck_volume(self):
        self.assertFalse(ducking.verify_keyframes(
            readback(0.0), "r2", [{"at": 20, "value": 1e-12}], 100.0))

    def test_tiny_positive_mismatch_is_not_hidden_by_absolute_tolerance(self):
        self.assertFalse(ducking.verify_keyframes(
            readback(1e-12), "r2", [{"at": 20, "value": 2e-12}], 100.0))

    def test_nonpositive_expected_volume_is_rejected(self):
        for value in (0, -1, -0.00001):
            with self.subTest(value=value):
                self.assertFalse(ducking.verify_keyframes(
                    readback(value), "r2", [{"at": 20, "value": value}], 100.0))

    def test_nonpositive_host_volume_is_rejected(self):
        for value in (0, -30):
            with self.subTest(value=value):
                self.assertFalse(ducking.verify_keyframes(
                    readback(value), "r2", [{"at": 20, "value": 30}], 100.0))

    def test_oversized_expected_curve_is_rejected(self):
        points = [{"at": i + 1, "value": 20.0} for i in range(2001)]
        self.assertFalse(ducking.verify_keyframes(readback(), "r2", points, 100.0))

    def test_oversized_host_animation_collection_is_rejected(self):
        report = readback()
        report["animations"].extend({"prop": "X"} for _ in range(256))
        self.assertFalse(ducking.verify_keyframes(
            report, "r2", [{"at": 20, "value": 30}], 100.0))

    def test_untrustworthy_revision_is_rejected(self):
        for revision in (" r2", "r2 ", "r2\n", "r2\x00", "x" * 257):
            with self.subTest(revision=revision):
                self.assertFalse(ducking.verify_keyframes(
                    readback(revision=revision), revision,
                    [{"at": 20, "value": 30}], 100.0))

    def test_large_finite_volume_still_verifies(self):
        report = readback(1e100)
        self.assertTrue(ducking.verify_keyframes(
            report, "r2", [{"at": 20, "value": 1e100}], 100.0))


class HostStateSafetyTests(unittest.IsolatedAsyncioTestCase):
    async def test_duplicate_voice_ids_fail_before_keyframe_read(self):
        items = {"items": [BGM, VOICE, {**VOICE, "frame": 45}]}
        for selection in (None, ["voice:1"]):
            args = {"action": "duck_bgm", "bgm_item_id": "bgm:1"}
            if selection is not None:
                args["voice_item_ids"] = selection
            with self.subTest(selection=selection), \
                 patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get, \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch(args)
            self.assertEqual(result["error_code"], "VOICE_ITEM_AMBIGUOUS")
            get.assert_awaited_once_with("/items")
            post.assert_not_awaited()

    async def test_malformed_animation_payload_fails_closed(self):
        for animations in (None, 7, {}, "Volume", [None] * 257):
            state = {"success": True, "revision": "r1", "animations": animations}
            with self.subTest(animations_type=type(animations).__name__), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, state])), \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch(
                    {"action": "duck_bgm", "bgm_item_id": "bgm:1"})
            self.assertEqual(result["error_code"], "BGM_VOLUME_UNSUPPORTED")
            post.assert_not_awaited()

    async def test_noninteger_baseline_frame_is_not_accepted(self):
        for frame in (False, 0.0, "0"):
            state = {"success": True, "revision": "r1", "animations": [
                {"prop": "Volume", "keyframes": [{"at": frame, "value": 100.0}]}]}
            with self.subTest(frame=frame), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, state])), \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch(
                    {"action": "duck_bgm", "bgm_item_id": "bgm:1"})
            self.assertEqual(result["error_code"], "BGM_HAS_EXISTING_KEYFRAMES")
            post.assert_not_awaited()

    async def test_unchanged_revision_after_set_is_unknown_outcome(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                 ITEMS, KEYS, {"success": True, "isSaved": True}, ITEMS])), \
             patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                 {"success": True, "backup_path": "C:/backup/cp.ymmp"},
                 {"success": True, "revision": "r1"}])) as post:
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        self.assertEqual(result["backup_path"], "C:/backup/cp.ymmp")
        self.assertEqual(len(post.await_args_list), 2)


    async def test_checkpoint_timeout_is_reported_without_writing_keyframes(self):
        with (
            patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                ITEMS, KEYS, {"success": True, "isSaved": True}])),
            patch.object(server, "ymm4_post", AsyncMock(
                side_effect=TimeoutError("checkpoint timeout"))) as post,
        ):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_BACKUP_OUTCOME_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        self.assertEqual(len(post.await_args_list), 1)

    async def test_snapshot_timeout_keeps_backup_path_and_prevents_edits(self):
        with (
            patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                ITEMS, KEYS, {"success": True, "isSaved": True}])),
            patch.object(server, "ymm4_post", AsyncMock(return_value={
                "success": True, "backup_path": "C:/backup/cp.ymmp"})) as post,
            patch.object(server.quality_guard, "verify_snapshot",
                         AsyncMock(side_effect=TimeoutError("snapshot timeout"))),
        ):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_DUCKING_SNAPSHOT_UNKNOWN")
        self.assertEqual(result["backup_path"], "C:/backup/cp.ymmp")
        self.assertEqual(result["applied"], [])
        self.assertEqual(len(post.await_args_list), 1)

    async def test_nonabsolute_or_padded_backup_path_blocks_mutation(self):
        for backup in ("relative/backup.ymmp", " C:/backup/cp.ymmp", "C:/backup/cp.ymmp "):
            with (
                self.subTest(backup=backup),
                patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                    ITEMS, KEYS, {"success": True, "isSaved": True}])),
                patch.object(server, "ymm4_post", AsyncMock(return_value={
                    "success": True, "backup_path": backup})) as post,
            ):
                result = await server.dispatch({
                    "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
            self.assertEqual(result["error_code"], "BGM_BACKUP_REQUIRED")
            self.assertEqual(len(post.await_args_list), 1)


if __name__ == "__main__":
    unittest.main()
