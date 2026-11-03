"""Fail-closed preflight for untrustworthy YMM4 ripple source snapshots."""
import unittest

from ripple_planning import plan_ripple_delete


def item(identity="a", frame=20, length=10, **fields):
    return {"item_id": identity, "revision": "r1", "frame": frame,
            "length": length, "layer": 1, **fields}


def plan(items, **snapshot_fields):
    return plan_ripple_delete(
        {"items": items, **snapshot_fields}, start_frame=10, end_frame=20,
        expected_revisions={"a": "r1"})


class RippleHostSnapshotIntegrityTests(unittest.TestCase):
    def test_untrustworthy_host_status_cannot_authorize_plan(self):
        for status in (False, None, 0, 1, "true", "false"):
            with self.subTest(status=status):
                self.assertEqual(
                    plan([item()], success=status)["error_code"],
                    "RIPPLE_SNAPSHOT_UNAVAILABLE")
        self.assertEqual(plan([item()], success=True)["move_count"], 1)
        self.assertEqual(plan([item()])["move_count"], 1)

    def test_partial_or_invalid_host_count_fails_closed(self):
        for count in (0, 2, True, "1", None, -1):
            with self.subTest(count=count):
                self.assertEqual(
                    plan([item()], success=True, count=count)["error_code"],
                    "RIPPLE_SNAPSHOT_UNAVAILABLE")
        self.assertEqual(plan([item()], success=True, count=1)["move_count"], 1)

    def test_end_frame_evidence_must_match_geometry(self):
        for key in ("end_frame", "endFrame"):
            for value in (29, 31, -1, None, "30", True):
                with self.subTest(field=key, value=value):
                    self.assertEqual(
                        plan([item(**{key: value})])["error_code"],
                        "RIPPLE_SNAPSHOT_INVALID")
            with self.subTest(field=key, value="valid"):
                self.assertEqual(plan([item(**{key: 30})])["move_count"], 1)
        self.assertEqual(
            plan([item(end_frame=30, endFrame=31)])["error_code"],
            "RIPPLE_SNAPSHOT_INVALID")
        self.assertEqual(
            plan([item(end_frame=30, endFrame=30)])["move_count"], 1)

    def test_ambiguous_host_item_ids_rejected(self):
        for identity in ("a\u202eb", "a\u200bb", "a\u2028b", "a\ufeffb",
                         "a\ufdd0b", "a\u200db", "e\u0301"):
            with self.subTest(identity=repr(identity)):
                self.assertEqual(
                    plan([item(identity)])["error_code"], "RIPPLE_ID_AMBIGUOUS")
        self.assertEqual(
            plan([item("\u00e9")], success=True)["error_code"],
            "RIPPLE_CONFLICT")  # valid ID, but no matching expected revision

    def test_unsafe_expected_revision_map_keys_rejected(self):
        with self.assertRaisesRegex(ValueError, "valid item IDs"):
            plan_ripple_delete(
                {"items": [item()]}, start_frame=10, end_frame=20,
                expected_revisions={"a\u200bb": "r1"})



class RippleReadbackStatusTests(unittest.TestCase):
    def test_bad_before_and_after_status_never_verify_edit(self):
        from ripple_verification import verify_ripple_result
        stable = {"item_id": "a", "frame": 0, "length": 10,
                  "layer": 1, "revision": "r1"}
        good = {"success": True, "count": 1, "items": [stable]}
        kwargs = {"start_frame": 10, "end_frame": 20}
        self.assertTrue(verify_ripple_result(good, good, **kwargs)["passed"])
        for bad in (False, None, 0, 1, "true", "false"):
            with self.subTest(status=bad):
                invalid = {**good, "success": bad}
                self.assertEqual(
                    verify_ripple_result(invalid, good, **kwargs)["error_code"],
                    "RIPPLE_BEFORE_SNAPSHOT_INVALID")
                self.assertEqual(
                    verify_ripple_result(good, invalid, **kwargs)["error_code"],
                    "RIPPLE_AFTER_SNAPSHOT_INVALID")



class RippleRevisionIntegrityTests(unittest.TestCase):
    def test_moved_item_needs_valid_readback_revision(self):
        from ripple_verification import verify_ripple_result
        before = {"success": True, "count": 2,
                  "items": [item("a", 20), item("b", 0)]}
        kwargs = {"start_frame": 10, "end_frame": 20,
                  "expected_revisions": {"a": "r1"}}
        good = {"success": True, "count": 2,
                "items": [item("a", 10), item("b", 0)]}
        self.assertTrue(verify_ripple_result(before, good, **kwargs)["passed"])
        for revision in (None, False, 0, "contains whitespace"):
            with self.subTest(revision=revision):
                bad = {"success": True, "count": 2,
                       "items": [{**item("a", 10), "revision": revision},
                                 item("b", 0)]}
                self.assertEqual(
                    verify_ripple_result(before, bad, **kwargs)["error_code"],
                    "RIPPLE_AFTER_SNAPSHOT_INVALID")
        no_revision = {"success": True, "count": 2,
                       "items": [{k: v for k, v in item("a", 10).items()
                                  if k != "revision"}, item("b", 0)]}
        result = verify_ripple_result(before, no_revision, **kwargs)
        self.assertFalse(result["passed"])
        self.assertIn("RIPPLE_REVISION_UNAVAILABLE",
                      [problem["code"] for problem in result["issues"]])

    def test_untouched_item_revision_drift_is_not_certified(self):
        from ripple_verification import verify_ripple_result
        before = {"success": True, "count": 2,
                  "items": [item("a", 20), item("b", 0)]}
        after = {"success": True, "count": 2,
                 "items": [item("a", 10), item("b", 0, revision="r2")]}
        result = verify_ripple_result(before, after, start_frame=10,
            end_frame=20, expected_revisions={"a": "r1"})
        self.assertFalse(result["passed"])
        self.assertIn("revision", [problem.get("field")
                                   for problem in result["issues"]])


if __name__ == "__main__":
    unittest.main()
