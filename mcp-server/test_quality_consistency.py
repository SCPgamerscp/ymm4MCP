"""Regressions for observed-state QA and voice-derived edits."""
import copy
import unittest
from unittest.mock import AsyncMock, patch

import editing
import scene_qa
import server
import visual_qa
from test_ducking import ITEMS, KEYS


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 100}]


class SubtitleTimingTests(unittest.TestCase):
    def items(self, spans):
        return [{"item_id": "v", "type": "VoiceItem", "text": "hello world",
                 "frame": 10, "length": 60, "layer": 1}, *[
            {"item_id": f"t{i}", "type": "TextItem", "text": "hello\nworld",
             "frame": start, "length": end - start, "layer": 9}
            for i, (start, end) in enumerate(spans)]]

    def test_one_frame_overlap_is_not_synced(self):
        result = editing.validate_timeline(self.items([(69, 80)]), subtitle_layers=[9])
        issue = result["issues"][0]
        self.assertEqual(issue["code"], "SUBTITLE_TIMING_MISMATCH")
        self.assertEqual(issue["uncovered_ranges"], [{"start_frame": 10, "end_frame": 69}])

    def test_split_subtitles_cover_voice_as_a_union(self):
        result = editing.validate_timeline(self.items([(5, 40), (40, 90)]), subtitle_layers=[9])
        self.assertTrue(result["passed"])

    def test_internal_gap_and_tolerance(self):
        items = self.items([(12, 40), (42, 68)])
        strict = editing.validate_timeline(items, subtitle_layers=[9])
        relaxed = editing.validate_timeline(items, subtitle_layers=[9], subtitle_timing_tolerance_frames=2)
        self.assertFalse(strict["passed"])
        self.assertEqual(len(strict["issues"][1]["uncovered_ranges"]), 3)  # layer GAP precedes subtitle issue
        self.assertTrue(relaxed["passed"])
        self.assertNotEqual(strict["criteria_hash"], relaxed["criteria_hash"])
        with self.assertRaises(ValueError):
            editing.validate_timeline(items, subtitle_timing_tolerance_frames=True)

    def test_wrong_text_cannot_fill_matching_subtitle_gap(self):
        items = self.items([(10, 30), (30, 70)])
        items[-1]["text"] = "different"
        result = editing.validate_timeline(items, subtitle_layers=[9])
        self.assertEqual(result["issues"][0]["code"], "SUBTITLE_TIMING_MISMATCH")


class SceneCoverageTests(unittest.TestCase):
    def test_audio_cannot_fill_visual_gap_and_edges_are_checked(self):
        items = [{"item_id": "a", "type": "AudioItem", "frame": 0, "length": 100},
                 {"item_id": "v", "type": "VideoItem", "frame": 10, "length": 80}]
        result = scene_qa.inspect(items, {"scene_ranges": RANGES})
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["code"], "SCENE_VISUAL_GAP")
        self.assertEqual(result["issues"][0]["uncovered_ranges"], [
            {"start_frame": 0, "end_frame": 10}, {"start_frame": 90, "end_frame": 100}])

    def test_tolerance_and_intentional_silent_scene(self):
        items = [{"item_id": "v", "type": "ImageItem", "frame": 10, "length": 80}]
        result = scene_qa.inspect(items, {"scene_ranges": RANGES, "require_audio": False,
                                        "max_visual_gap_frames": 10})
        self.assertTrue(result["passed"])

    def test_unknown_types_do_not_provide_coverage(self):
        result = scene_qa.inspect([{"item_id": "x", "type": "UnknownItem", "frame": 0, "length": 100}],
                                  {"scene_ranges": RANGES})
        self.assertEqual(len(result["issues"]), 2)

    def test_invalid_configs(self):
        for config in ({}, {"scene_ranges": RANGES, "require_audio": "false"},
                       {"scene_ranges": RANGES, "require_visual": False, "require_audio": False},
                       {"scene_ranges": RANGES, "max_audio_gap_frames": -1},
                       {"scene_ranges": RANGES, "unexpected": 1}):
            with self.subTest(config=config), self.assertRaises(ValueError):
                scene_qa.inspect([], config)


class QualityConsistencyTests(unittest.IsolatedAsyncioTestCase):
    async def test_scene_gaps_block_gate_and_criteria_are_bound(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": []})):
            args = {"action": "qa_gate", "scene_check": {"scene_ranges": RANGES}}
            first = await server.dispatch(args)
            self.assertEqual(first["decision"], "repair")
            self.assertEqual(first["qa"]["issues"][0]["source"], "scene")
            with self.assertRaisesRegex(ValueError, "same validation criteria"):
                await server.dispatch({**args, "scene_check": {"scene_ranges": RANGES, "require_audio": False},
                                       "qa_history": [first["qa"]]})

    async def test_changed_timeline_invalidates_clean_visual_result(self):
        before = {"items": [{"item_id": "x", "frame": 0, "length": 10, "layer": 1, "revision": "r1"}]}
        after = copy.deepcopy(before)
        after["items"][0]["revision"] = "r2"
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[before, after])), \
             patch.object(server, "run_visual_qa", AsyncMock(return_value={
                 "success": True, "passed": True, "issues": []})):
            result = await server.dispatch({"action": "qa_gate", "visual_check": {"end_frame": 0}})
        self.assertEqual(result["error_code"], "QA_SNAPSHOT_CONFLICT")
        self.assertFalse(result["passed"])
        self.assertNotIn("qa", result)

    async def test_reordered_snapshot_is_accepted(self):
        items = [{"item_id": "a", "frame": 0, "length": 10, "layer": 1},
                 {"item_id": "b", "frame": 0, "length": 10, "layer": 2}]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"items": items}, {"items": list(reversed(items))}])), \
             patch.object(server, "run_visual_qa", AsyncMock(return_value={
                 "success": True, "passed": True, "issues": []})):
            result = await server.dispatch({"action": "qa_gate", "visual_check": {"end_frame": 0}})
        self.assertEqual(result["decision"], "pass")
        self.assertEqual(len(result["qa"]["snapshot_hash"]), 64)

    async def test_snapshot_read_failure_does_not_accept(self):
        for bad in ({"success": False, "error": "unavailable"}, {"items": [None]}, TimeoutError("read failed")):
            with self.subTest(bad=bad), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=[{"items": []}, bad])), \
                 patch.object(server, "run_visual_qa", AsyncMock(return_value={
                     "success": True, "passed": True, "issues": []})):
                result = await server.dispatch({"action": "qa_gate", "visual_check": {"end_frame": 0}})
                self.assertFalse(result["passed"])
                self.assertNotIn("decision", result)

    def test_different_missing_expected_item_is_not_stalled(self):
        expected = [{"item_id": "a"}, {"item_id": "b"}]
        def report(identity):
            return editing.validate_timeline([{"item_id": identity, "frame": 0, "length": 10, "layer": 1}],
                                              expected=expected)
        self.assertEqual(editing.evaluate_qa_gate(report("b"), [report("a")])["reason_code"], "QA_ISSUES_REMAIN")
        self.assertEqual(editing.evaluate_qa_gate(report("a"), [report("a")])["reason_code"], "QA_STALLED")

    def test_no_preview_samples_cannot_pass(self):
        with self.assertRaisesRegex(ValueError, "1..40"):
            visual_qa.inspect([], step_frames=30)


class DuckingConsistencyTests(unittest.IsolatedAsyncioTestCase):
    async def test_voice_change_after_backup_stops_before_first_keyframe(self):
        changed = copy.deepcopy(ITEMS)
        changed["items"][1]["frame"] = 50
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                 ITEMS, KEYS, {"success": True, "isSaved": True}, changed])), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={
                 "success": True, "backup_path": "C:/backup.ymmp"})) as post:
            result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
        self.assertEqual(result["error_code"], "BGM_DUCKING_SNAPSHOT_CONFLICT")
        self.assertEqual(result["applied"], [])
        self.assertEqual(post.await_count, 1)

    async def test_readback_requires_exact_revision_and_key_set(self):
        points = [{"at": 0, "value": 100}, {"at": 15, "value": 100}, {"at": 20, "value": 30},
                  {"at": 30, "value": 30}, {"at": 40, "value": 100}]
        base = {"success": True, "revision": "r5", "animations": [{"prop": "Volume", "keyframes": points}]}
        variants = []
        stale = copy.deepcopy(base)
        stale["revision"] = "r6"
        variants.append(stale)
        for extra in ({"at": 21, "value": 100}, {"at": 20, "value": 30}):
            altered = copy.deepcopy(base)
            altered["animations"][0]["keyframes"].append(extra)
            variants.append(altered)
        for bad in (True, float("nan")):
            altered = copy.deepcopy(base)
            altered["animations"][0]["keyframes"][0]["value"] = bad
            variants.append(altered)
        for readback in variants:
            with self.subTest(readback=readback), \
                 patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                     ITEMS, KEYS, {"success": True, "isSaved": True}, ITEMS, readback])), \
                 patch.object(server, "ymm4_post", AsyncMock(side_effect=[
                     {"success": True, "backup_path": "C:/backup.ymmp"}, *[
                         {"success": True, "revision": f"r{i}"} for i in range(2, 6)]])):
                result = await server.dispatch({"action": "duck_bgm", "bgm_item_id": "bgm:1", "dry_run": False})
                self.assertEqual(result["error_code"], "BGM_DUCKING_VERIFY_FAILED")
                self.assertEqual(result["backup_path"], "C:/backup.ymmp")


if __name__ == "__main__":
    unittest.main()
