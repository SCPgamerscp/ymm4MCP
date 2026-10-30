"""Regression tests for server-owned, short-lived ripple pre-edit evidence."""
import unittest

from ripple_evidence import RippleEvidenceStore


def snapshot(*items):
    return {"success": True, "count": len(items), "items": list(items)}


BEFORE = snapshot(
    {"item_id": "cut", "frame": 10, "length": 10, "layer": 1,
     "revision": "r1", "effects": ["fade"]},
    {"item_id": "later", "frame": 20, "length": 10, "layer": 1,
     "revision": "r2", "effects": ["glow"]})
AFTER = snapshot(
    {"item_id": "later", "frame": 10, "length": 10, "layer": 1,
     "revision": "r3", "effects": ["glow"]})


class EvidenceTests(unittest.TestCase):
    def setUp(self):
        self.now = [100.0]
        self.store = RippleEvidenceStore(clock=lambda: self.now[0])

    def capture(self, before=BEFORE):
        return self.store.capture(
            before, start_frame=10, end_frame=20, layers=[1],
            expected_revisions={"cut": "r1", "later": "r2"})

    def test_verified_capture_and_immutable_original(self):
        before = snapshot(*(dict(item) for item in BEFORE["items"]))
        evidence = self.capture(before)
        before["items"][0]["frame"] = 123
        result = self.store.verify(self.store.lookup(evidence["plan_id"]), AFTER)
        self.assertTrue(result["passed"])
        self.assertFalse(result["atomicity_verified"])
        self.assertEqual(result["before_sha256"], evidence["before_sha256"])

    def test_nested_content_change_is_rejected(self):
        evidence = self.capture()
        altered = snapshot(dict(AFTER["items"][0], effects=["different"]))
        result = self.store.verify(self.store.lookup(evidence["plan_id"]), altered)
        self.assertFalse(result["passed"])
        self.assertTrue(any(x["code"] == "RIPPLE_CONTENT_CHANGED"
                            for x in result["issues"]))

    def test_expiration_and_unknown_id(self):
        evidence = self.capture()
        self.assertIsNone(self.store.lookup("forged"))
        self.now[0] += self.store.TTL_SECONDS
        self.assertIsNone(self.store.lookup(evidence["plan_id"]))

    def test_expiration_during_host_fetch(self):
        evidence = self.capture()
        record = self.store.lookup(evidence["plan_id"])
        self.now[0] += self.store.TTL_SECONDS
        self.assertEqual(self.store.verify(record, AFTER)["error_code"],
                         "RIPPLE_EVIDENCE_EXPIRED")

    def test_bounded_store_evicts_oldest(self):
        first = self.capture()["plan_id"]
        for _ in range(self.store.MAX_RECORDS):
            self.capture()
        self.assertIsNone(self.store.lookup(first))

    def test_reject_incomplete_inventory(self):
        with self.assertRaises(ValueError):
            self.capture({"items": BEFORE["items"]})

    def test_reject_oversized_inventory(self):
        self.store.MAX_BYTES = 100
        with self.assertRaises(ValueError):
            self.capture()

    def test_no_unsafe_replay(self):
        evidence = self.capture()
        result = self.store.verify(self.store.lookup(evidence["plan_id"]), AFTER)
        self.assertNotIn("operations", result)
        self.assertFalse(result["executed"])


if __name__ == "__main__":
    unittest.main()
