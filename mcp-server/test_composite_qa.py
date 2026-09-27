"""The quality gate must include actual visual/audio results and reject stale history."""
import unittest
from unittest.mock import AsyncMock, patch

import editing
import server


class CompositeQaTests(unittest.IsolatedAsyncioTestCase):
    async def test_audio_error_blocks_structurally_clean_timeline(self):
        async def get(path):
            if path == "/items":
                return {"success": True, "items": []}
            return {"success": True, "passed": False, "issues": [
                {"code": "LONG_SILENCE", "severity": "error", "startSeconds": 0, "endSeconds": 3}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)):
            result = await server.dispatch({"action": "qa_gate", "audio_check": {"path": "C:/audio.wav"}})
        self.assertEqual(result["decision"], "repair")
        self.assertEqual(result["qa"]["issues"][0]["source"], "audio")
        self.assertFalse(result["qa"]["passed"])

    async def test_visual_failure_does_not_pass_or_pollute_history(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": []})), \
             patch.object(server, "run_visual_qa", AsyncMock(return_value={
                 "success": False, "passed": False, "error": "capture failed"})):
            result = await server.dispatch({"action": "qa_gate", "visual_check": {"end_frame": 30}})
        self.assertEqual(result["error_code"], "QA_CHECK_FAILED")
        self.assertFalse(result["passed"])

    async def test_gate_detects_improvement_and_checks_same_criteria(self):
        bad = {"success": True, "passed": False, "issues": [
            {"code": "BLACK_FRAME", "severity": "error", "start_frame": 0, "end_frame": 30}]}
        clean = {"success": True, "passed": True, "issues": []}
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"items": []})), \
             patch.object(server, "run_visual_qa", AsyncMock(side_effect=[bad, clean, clean])):
            args = {"action": "qa_gate", "visual_check": {"end_frame": 30, "black_as_error": True}}
            first = await server.dispatch(args)
            self.assertEqual(first["qa"]["issues"][0]["frame_range"], [0, 30])
            second = await server.dispatch({**args, "qa_history": [first["qa"]]})
            self.assertEqual((second["decision"], second["score_delta"]), ("pass", 20))
            with self.assertRaisesRegex(ValueError, "same validation criteria"):
                await server.dispatch({**args, "visual_check": {"end_frame": 60, "black_as_error": True},
                                       "qa_history": [first["qa"]]})

    def test_inconsistent_external_report_is_rejected(self):
        structural = editing.validate_timeline([])
        with self.assertRaisesRegex(ValueError, "failed without an error"):
            editing.combine_qa_reports(structural, {"visual": {
                "success": True, "passed": False, "issues": []}}, {"visual": {"end_frame": 30}})

    def test_moved_silence_does_not_look_stalled(self):
        structural = editing.validate_timeline([])
        criteria = {"audio": {"path": "C:/audio.wav"}}
        def report(start):
            return editing.combine_qa_reports(structural, {"audio": {
                "success": True, "passed": False, "issues": [
                    {"code": "LONG_SILENCE", "severity": "error",
                     "startSeconds": start, "endSeconds": start + 3}]}}, criteria)
        self.assertEqual(editing.evaluate_qa_gate(report(10), [report(0)])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_qa_stall_ignores_order_of_affected_item_ids(self):
        structural = editing.validate_timeline([])
        def report(item_ids):
            return editing.combine_qa_reports(structural, {"visual": {
                "success": True, "passed": False, "issues": [
                    {"code": "OVERLAP", "severity": "error", "item_ids": item_ids}]}},
                {"visual": {"end_frame": 30}})
        decision = editing.evaluate_qa_gate(report(["b", "a"]), [report(["a", "b"])])
        self.assertEqual(decision["reason_code"], "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
