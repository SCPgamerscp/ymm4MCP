"""Additional BGM ducking trust-boundary regressions (#61)."""
import unittest
from unittest.mock import AsyncMock, patch

import ducking
import server


BGM = {"item_id": "bgm:1", "revision": "r1", "type": "AudioItem",
       "frame": 0, "length": 100}
VOICE = {"item_id": "voice:1", "revision": "v1", "type": "VoiceItem",
         "frame": 20, "length": 10}
ITEMS = {"items": [BGM, VOICE]}
KEYS = {"success": True, "revision": "r1", "animations": [
    {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}]}]}


class IdentifierValidationTests(unittest.TestCase):
    def test_reject_ambiguous_or_unsafe_host_identifiers(self):
        for value in (None, "", " x", "x ", "x\x00", "x\n", "x\x7f",
                      "x\x85", "x\ud800", "x" * 1025):
            with self.subTest(value=repr(value)):
                self.assertFalse(ducking.valid_identifier(value))

    def test_accept_normal_stable_ids(self):
        for value in ("native:abc-123", "runtime:0123", "voice:日本語"):
            with self.subTest(value=value):
                self.assertTrue(ducking.valid_identifier(value))

    def test_revision_limit_is_stricter_than_item_id_limit(self):
        self.assertTrue(ducking.valid_identifier("r" * 257))
        self.assertFalse(ducking.valid_identifier("r" * 257, max_length=256))


class HostBoundaryTests(unittest.IsolatedAsyncioTestCase):
    async def test_malformed_bgm_id_fails_before_http(self):
        for item_id in (" bgm:1", "bgm:1 ", "bgm:1\n", "x" * 1025):
            with (self.subTest(item_id=repr(item_id)),
                  patch.object(server, "ymm4_get", new_callable=AsyncMock) as get):
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "duck_bgm", "bgm_item_id": item_id})
                get.assert_not_awaited()

    async def test_invalid_snapshot_revision_fails_before_keyframe_query(self):
        for revision in (" r1", "r1\x00", "r" * 257):
            items = {"items": [{**BGM, "revision": revision}, VOICE]}
            with (self.subTest(revision=repr(revision)),
                  patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get):
                result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
                self.assertEqual(result["error_code"], "BGM_REVISION_UNAVAILABLE")
                self.assertEqual(get.await_count, 1)

    async def test_stale_expected_revision_prevents_keyframe_reads(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "expected_revision": "old-r1"})
        self.assertEqual(result["error_code"], "REVISION_CONFLICT")
        self.assertEqual(get.await_count, 1)

    async def test_malformed_expected_revision_is_rejected(self):
        for expected in (" r1", "r1\\x00", "x" * 257):
            with (self.subTest(expected=repr(expected)),
                  patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get):
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                           "expected_revision": expected})
                self.assertEqual(get.await_count, 1)

    async def test_matching_expected_revision_still_plans(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, KEYS])):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "expected_revision": "r1"})
        self.assertTrue(result["success"])

    async def test_invalid_host_voice_id_blocks_planning(self):
        for bad_id in (" voice:1", "voice:1\x00", "x" * 1025):
            items = {"items": [BGM, {**VOICE, "item_id": bad_id}]}
            with (self.subTest(bad_id=repr(bad_id)),
                  patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get):
                result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
                self.assertEqual(result["error_code"], "VOICE_ITEM_ID_INVALID")
                self.assertEqual(get.await_count, 1)

    async def test_invalid_explicit_voice_selection_rejected(self):
        for bad_id in ("voice:1 ", "voice:1\x00", "x" * 1025):
            with (self.subTest(bad_id=repr(bad_id)),
                  patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get):
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                           "voice_item_ids": [bad_id]})
                self.assertEqual(get.await_count, 1)

    async def test_bgm_id_colliding_with_non_audio_item_is_ambiguous(self):
        items = {"items": [BGM, VOICE, {"item_id": "bgm:1", "type": "TextItem"}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
        self.assertEqual(result["error_code"], "ITEM_ID_AMBIGUOUS")
        self.assertEqual(get.await_count, 1)

    async def test_selected_voice_id_colliding_with_text_item_is_ambiguous(self):
        items = {"items": [BGM, VOICE, {"item_id": "voice:1", "type": "TextItem"}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "voice_item_ids": ["voice:1"]})
        self.assertEqual(result["error_code"], "ITEM_ID_AMBIGUOUS")
        self.assertEqual(get.await_count, 1)

    async def test_invalid_host_baseline_never_authorizes_ducking(self):
        for baseline in ("100", True, float("nan"), float("inf"), -1, 0, 10 ** 1000):
            keys = {"success": True, "revision": "r1", "animations": [
                {"prop": "Volume", "keyframes": [{"at": 0, "value": baseline}]}]}
            with (self.subTest(baseline=repr(baseline)[:40]),
                  patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, keys])),
                  patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
                result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
                self.assertEqual(result["error_code"], "BGM_VOLUME_UNSUPPORTED")
                post.assert_not_awaited()

    async def test_normal_baseline_still_plans(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, KEYS])):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
        self.assertTrue(result["success"])
        self.assertTrue(result["dry_run"])
        self.assertGreater(len(result["keyframes"]), 1)

    async def test_backup_must_be_a_bounded_project_file(self):
        for backup in ("C:/backup/not-a-project.txt", "C:/backup/file.ymmp:stream",
                       "C:/backup/" + "x" * 4096 + ".ymmp"):
            with (self.subTest(backup=backup[:50]),
                  patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                      ITEMS, KEYS, {"success": True, "isSaved": True}])),
                  patch.object(server, "ymm4_post", AsyncMock(return_value={
                      "success": True, "backup_path": backup})) as post):
                result = await server.dispatch({"action": "duck_bgm",
                                                "bgm_item_id": "bgm:1", "dry_run": False})
                self.assertEqual(result["error_code"], "BGM_BACKUP_REQUIRED")
                self.assertEqual(post.await_count, 1)

    async def _attempt_mutation(self, response):
        with (patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                  ITEMS, KEYS, {"success": True, "isSaved": True}])),
              patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                  {"success": True, "backup_path": "C:/backup/cp.ymmp"}, response])) as post,
              patch.object(server.quality_guard, "verify_snapshot",
                           AsyncMock(return_value=None))):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1",
                                            "dry_run": False})
        self.assertEqual(post.await_count, 2)
        return result

    async def test_malformed_revision_after_mutation_is_unknown(self):
        for revision in ("r2 ", "r2\x00", "r" * 257):
            with self.subTest(revision=repr(revision)):
                result = await self._attempt_mutation(
                    {"success": True, "revision": revision})
                self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
                self.assertTrue(result["outcome_unknown"])
                self.assertEqual(result["applied"], [])
                self.assertEqual(result["backup_path"], "C:/backup/cp.ymmp")

    async def test_rejected_host_mutation_is_unknown_not_safe_retry(self):
        for response in ({"success": False, "revision": "r2"},
                         {"success": False, "error_code": "HOST_FAILED"}, None):
            with self.subTest(response=response):
                result = await self._attempt_mutation(response)
                self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
                self.assertTrue(result["outcome_unknown"])
                self.assertEqual(result["applied"], [])


if __name__ == "__main__":
    unittest.main()
