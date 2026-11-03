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


if __name__ == "__main__":
    unittest.main()
