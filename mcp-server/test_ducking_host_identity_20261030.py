"""BGM ducking item identity, host echo, and revision-chain regressions (#61)."""
import unittest
from unittest.mock import AsyncMock, patch

import ducking
import server


BGM = {"item_id": "bgm:1", "revision": "r1", "type": "AudioItem",
       "frame": 0, "length": 100}
VOICE = {"item_id": "voice:1", "revision": "v1", "type": "VoiceItem",
         "frame": 20, "length": 10}
ITEMS = {"items": [BGM, VOICE]}
KEYS = {"success": True, "revision": "r1", "item_id": "bgm:1",
        "animations": [{"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}]}]}
BACKUP = {"success": True, "backup_path": "C:/backup/cp.ymmp"}


def readback(item_id="bgm:1"):
    points = ducking.plan(BGM, [VOICE], 100.0)
    report = {"success": True, "revision": "r5", "animations": [
        {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0}, *points]}]}
    if item_id is not None:
        report["item_id"] = item_id
    return report


class ReadbackItemIdentityTests(unittest.TestCase):
    def test_correct_item_identity_verifies(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        self.assertTrue(ducking.verify_keyframes(
            readback(), "r5", points, 100.0, expected_item_id="bgm:1"))

    def test_wrong_item_identity_never_verifies(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        self.assertFalse(ducking.verify_keyframes(
            readback("bgm:other"), "r5", points, 100.0, expected_item_id="bgm:1"))

    def test_missing_item_identity_never_verifies(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        self.assertFalse(ducking.verify_keyframes(
            readback(None), "r5", points, 100.0, expected_item_id="bgm:1"))

    def test_invalid_expected_identity_is_rejected(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        for value in (" bgm:1", "bgm:1\x00", "x" * 1025):
            with self.subTest(value=repr(value)[:30]):
                self.assertFalse(ducking.verify_keyframes(
                    readback(), "r5", points, 100.0, expected_item_id=value))

    def test_per_write_echo_requires_matching_value(self):
        point = {"at": 15, "value": 100.0}
        self.assertFalse(ducking.verify_mutation_point(
            [{"at": 15, "value": 10.0}], point))
        self.assertFalse(ducking.verify_mutation_point(
            [{"at": 15, "value": float("nan")}], point))
        self.assertFalse(ducking.verify_mutation_point(
            [{"at": 15, "value": 100.0}, {"at": 15, "value": 100.0}], point))
        self.assertTrue(ducking.verify_mutation_point(
            [{"at": 15, "value": 100.0}], point))

    def test_final_absolute_frame_mismatch_is_rejected(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        report = readback()
        report["animations"][0]["keyframes"][1]["absolute_frame"] = 999
        self.assertFalse(ducking.verify_keyframes(
            report, "r5", points, 100.0, expected_item_id="bgm:1",
            expected_start_frame=0))

    def test_correct_absolute_frame_metadata_verifies(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        report = readback()
        for key in report["animations"][0]["keyframes"]:
            key["absolute_frame"] = key["at"]
        self.assertTrue(ducking.verify_keyframes(
            report, "r5", points, 100.0, expected_item_id="bgm:1",
            expected_start_frame=0))


class DuckingHostIdentityTests(unittest.IsolatedAsyncioTestCase):
    async def test_oversized_snapshot_fails_before_keyframe_query(self):
        items = {"items": [BGM, VOICE] + [
            {"item_id": f"text:{i}", "type": "TextItem"} for i in range(9999)]}
        with (patch.object(server, "ymm4_get", AsyncMock(return_value=items)) as get,
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1"})
        self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_initial_keyframe_response_for_other_item_fails_closed(self):
        wrong = {**KEYS, "item_id": "bgm:other"}
        with (patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, wrong])),
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_KEYFRAMES_UNAVAILABLE")
        post.assert_not_awaited()


    async def test_initial_keyframe_wrong_timeline_frame_is_rejected(self):
        wrong = {**KEYS, "frame": 7}
        with (patch.object(server, "ymm4_get", AsyncMock(side_effect=[ITEMS, wrong])),
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_KEYFRAMES_UNAVAILABLE")
        post.assert_not_awaited()

    async def test_initial_keyframe_wrong_layer_is_rejected(self):
        items = {"items": [{**BGM, "layer": 3}, VOICE]}
        wrong = {**KEYS, "layer": 4}
        with (patch.object(server, "ymm4_get", AsyncMock(side_effect=[items, wrong])),
              patch.object(server, "ymm4_post", new_callable=AsyncMock) as post):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_KEYFRAMES_UNAVAILABLE")
        post.assert_not_awaited()

    async def _apply(self, mutation_responses, *, final_readback=None):
        gets = [ITEMS, KEYS, {"success": True, "isSaved": True}]
        if final_readback is not None:
            gets.append(final_readback)
        with (patch.object(server, "ymm4_get", AsyncMock(side_effect=gets)),
              patch.object(server, "ymm4_post",
                           AsyncMock(side_effect=[BACKUP, *mutation_responses])) as post,
              patch.object(server.quality_guard, "verify_snapshot",
                           AsyncMock(return_value=None))):
            result = await server.dispatch({
                "action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        return result, post.await_args_list

    async def test_mutation_echo_for_other_item_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2", "item_id": "bgm:other"}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertTrue(result["outcome_unknown"])
        self.assertEqual(result["backup_path"], BACKUP["backup_path"])
        self.assertEqual(result["applied"], [])
        self.assertEqual(len(calls), 2)

    async def test_mutation_echo_wrong_property_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2", "prop": "Opacity"}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(len(calls), 2)

    async def test_mutation_echo_wrong_action_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2", "action": "remove"}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(len(calls), 2)

    async def test_mutation_echo_wrong_frame_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2", "at": 99}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(len(calls), 2)


    async def test_mutation_echo_wrong_volume_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2",
             "keyframes": [{"at": 15, "value": 12.0}]}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(result["applied"], [])
        self.assertEqual(len(calls), 2)

    async def test_mutation_echo_duplicate_frame_stops_chain(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2",
             "keyframes": [{"at": 15, "value": 100.0},
                           {"at": 15, "value": 100.0}]}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(len(calls), 2)

    async def test_revision_cycle_stops_before_third_write(self):
        result, calls = await self._apply([
            {"success": True, "revision": "r2"},
            {"success": True, "revision": "r1"}])
        self.assertEqual(result["error_code"], "BGM_DUCKING_OUTCOME_UNKNOWN")
        self.assertEqual(len(result["applied"]), 1)
        self.assertEqual(len(calls), 3)

    async def test_valid_echo_and_unique_revisions_still_verify(self):
        points = ducking.plan(BGM, [VOICE], 100.0)
        replies = [
            {"success": True, "revision": f"r{i + 2}", "item_id": "bgm:1",
             "prop": "Volume", "action": "set", "at": point["at"],
             "keyframes": [{"at": point["at"], "value": point["value"]}]}
            for i, point in enumerate(points)]
        result, calls = await self._apply(replies, final_readback=readback())
        self.assertTrue(result["success"])
        self.assertTrue(result["verified"])
        self.assertEqual(len(calls), len(points) + 1)

    async def test_wrong_item_final_readback_never_certifies_success(self):
        replies = [{"success": True, "revision": f"r{i}"} for i in range(2, 6)]
        result, calls = await self._apply(
            replies, final_readback=readback("bgm:other"))
        self.assertEqual(result["error_code"], "BGM_DUCKING_VERIFY_FAILED")
        self.assertNotIn("verified", result)
        self.assertEqual(len(calls), 5)

    async def test_missing_item_final_readback_never_certifies_success(self):
        replies = [{"success": True, "revision": f"r{i}"} for i in range(2, 6)]
        result, calls = await self._apply(replies, final_readback=readback(None))
        self.assertEqual(result["error_code"], "BGM_DUCKING_VERIFY_FAILED")
        self.assertEqual(len(calls), 5)


if __name__ == "__main__":
    unittest.main()
