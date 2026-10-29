"""Regression tests for safe repair targeting and QA report integrity (Issue #61)."""
import copy
import unittest

import editing


def item(identity="unique", frame=15, length=10, layer=1):
    return {"item_id": identity, "frame": frame, "length": length, "layer": layer}


class SafeDurationRepairTests(unittest.TestCase):
    def test_unique_id_can_be_trimmed(self):
        report = editing.validate_timeline([item()], duration=20)
        fix = next(p for p in report["issues"] if p["code"] == "EXCEEDS_DURATION")["suggested_fix"]
        self.assertEqual((fix["item_id"], fix["value"]), ("unique", 5))

    def test_missing_id_has_no_unsafe_fix(self):
        entry = item()
        del entry["item_id"]
        report = editing.validate_timeline([entry], duration=20)
        issue = next(p for p in report["issues"] if p["code"] == "EXCEEDS_DURATION")
        self.assertNotIn("suggested_fix", issue)
        self.assertIn("unique", issue["message"])

    def test_invalid_id_has_no_unsafe_fix(self):
        for identity in ("", "  ", "bad\nname", 42):
            with self.subTest(identity=identity):
                report = editing.validate_timeline([item(identity)], duration=20)
                issue = next(p for p in report["issues"] if p["code"] == "EXCEEDS_DURATION")
                self.assertNotIn("suggested_fix", issue)

    def test_later_duplicate_disables_earlier_repair(self):
        report = editing.validate_timeline(
            [item("shared", layer=1), item("shared", frame=0, length=10, layer=2)],
            duration=20)
        issue = next(p for p in report["issues"] if p["code"] == "EXCEEDS_DURATION")
        self.assertNotIn("suggested_fix", issue)
        self.assertIn("DUPLICATE_ITEM_ID", [p["code"] for p in report["issues"]])

    def test_earlier_duplicate_disables_later_repair(self):
        report = editing.validate_timeline(
            [item("shared", frame=0, length=10, layer=2), item("shared", layer=1)],
            duration=20)
        issue = next(p for p in report["issues"] if p["code"] == "EXCEEDS_DURATION")
        self.assertNotIn("suggested_fix", issue)


class ExpectedSelectorSafetyTests(unittest.TestCase):
    def test_control_characters_in_stable_id_selector_rejected(self):
        for identity in ("bad\nname", "bad\x00name", "bad\x7fname"):
            with self.subTest(identity=identity):
                with self.assertRaisesRegex(ValueError, "valid string"):
                    editing.validate_timeline([item(identity)], expected=[{"item_id": identity}])

    def test_control_characters_in_legacy_alias_rejected(self):
        with self.assertRaisesRegex(ValueError, "valid string"):
            editing.validate_timeline([item("bad\tname")], expected=[{"id": "bad\tname"}])

    def test_invalid_geometry_cannot_satisfy_expected_selector(self):
        invalid = item("bad", frame=editing.MAX_FRAME, length=2)
        report = editing.validate_timeline([invalid], expected=[{"item_id": "bad"}])
        self.assertIn("INVALID_ITEM", [p["code"] for p in report["issues"]])
        self.assertIn("EXPECTED_NOT_FOUND", [p["code"] for p in report["issues"]])

    def test_invalid_geometry_does_not_make_valid_match_ambiguous(self):
        report = editing.validate_timeline(
            [item("same", frame=editing.MAX_FRAME, length=2),
             item("same", frame=0, length=10, layer=2)],
            expected=[{"item_id": "same"}], include_gaps=False)
        self.assertNotIn("EXPECTED_AMBIGUOUS", [p["code"] for p in report["issues"]])

    def test_valid_expected_selector_still_matches(self):
        report = editing.validate_timeline([item()], expected=[{"item_id": "unique"}])
        self.assertTrue(report["passed"], report["issues"])


class GateReportIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.clean = editing.validate_timeline([])

    def test_forged_valid_flag_rejected_in_current(self):
        forged = copy.deepcopy(self.clean)
        forged["valid"] = False
        with self.assertRaisesRegex(ValueError, "valid status"):
            editing.evaluate_qa_gate(forged)

    def test_forged_valid_flag_rejected_in_history(self):
        forged = copy.deepcopy(self.clean)
        forged["valid"] = False
        with self.assertRaisesRegex(ValueError, "valid status"):
            editing.evaluate_qa_gate(self.clean, [forged])

    def test_forged_issue_count_rejected(self):
        forged = copy.deepcopy(self.clean)
        forged["issue_count"] = 1
        with self.assertRaisesRegex(ValueError, "issue_count"):
            editing.evaluate_qa_gate(forged)

    def test_boolean_issue_count_rejected(self):
        forged = copy.deepcopy(self.clean)
        forged["issue_count"] = False
        with self.assertRaisesRegex(ValueError, "issue_count"):
            editing.evaluate_qa_gate(forged)

    def test_forged_summary_rejected(self):
        forged = copy.deepcopy(self.clean)
        forged["summary"]["warnings"] = 1
        with self.assertRaisesRegex(ValueError, "summary"):
            editing.evaluate_qa_gate(forged)

    def test_boolean_summary_counter_rejected(self):
        forged = copy.deepcopy(self.clean)
        forged["summary"]["errors"] = False
        with self.assertRaisesRegex(ValueError, "summary"):
            editing.evaluate_qa_gate(forged)

    def test_control_character_in_issue_reference_rejected(self):
        report = editing.validate_timeline(
            [item("a", frame=0), item("b", frame=5)])
        report["issues"][0]["item_ids"][0] = "bad\nname"
        with self.assertRaisesRegex(ValueError, "item_ids"):
            editing.evaluate_qa_gate(report)

    def test_blank_criteria_hash_rejected(self):
        forged = copy.deepcopy(self.clean)
        forged["criteria_hash"] = "  "
        with self.assertRaisesRegex(ValueError, "same validation criteria"):
            editing.evaluate_qa_gate(forged)


class CompositeIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.structural = editing.validate_timeline([])

    def test_forged_structural_summary_cannot_be_laundered(self):
        forged = copy.deepcopy(self.structural)
        forged["summary"]["errors"] = 5
        with self.assertRaisesRegex(ValueError, "summary"):
            editing.combine_qa_reports(forged, {}, {})

    def test_forged_structural_pass_status_cannot_be_laundered(self):
        forged = copy.deepcopy(self.structural)
        forged["passed"] = False
        with self.assertRaisesRegex(ValueError, "passed status"):
            editing.combine_qa_reports(forged, {}, {})

    def test_invalid_source_name_rejected(self):
        for source in ("", "bad\nsource", "a" * 65):
            with self.subTest(source=source):
                with self.assertRaisesRegex(ValueError, "source"):
                    editing.combine_qa_reports(
                        self.structural, {source: {"success": True, "passed": True, "issues": []}}, {})

    def test_external_invalid_item_ids_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid QA issue"):
            editing.combine_qa_reports(self.structural, {
                "visual": {"success": True, "passed": False, "issues": [
                    {"code": "OUTSIDE", "severity": "error", "item_ids": ["bad\x00id"]}]}}, {})

    def test_valid_composite_still_passes(self):
        result = editing.combine_qa_reports(self.structural, {
            "visual": {"success": True, "passed": True, "issues": []}}, {"visual": {"end_frame": 30}})
        self.assertEqual(editing.evaluate_qa_gate(result)["decision"], "pass")


if __name__ == "__main__":
    unittest.main()
