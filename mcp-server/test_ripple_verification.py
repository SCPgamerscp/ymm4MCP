"""Post-edit ripple acceptance tests (#61); no YMM4 host is mutated."""
import unittest
from ripple_verification import verify_ripple_result as verify


def item(name, frame, length=10, layer=1, revision="r1", **extra):
    return {"item_id": name, "frame": frame, "length": length,
            "layer": layer, "revision": revision, **extra}


def snap(*items):
    return {"success": True, "count": len(items), "items": list(items)}


class RippleVerificationTests(unittest.TestCase):
    def setUp(self):
        self.before = snap(item("left", 0, type="TextItem", text="hello"),
                           item("cut", 10, type="AudioItem", source_path="C:/a.wav"),
                           item("right", 20, type="VideoItem", source_path="C:/b.mp4"),
                           item("other", 20, layer=2, type="TextItem", text="keep"))
        self.revisions = {"cut": "r1", "right": "r1"}
        self.after = snap(item("left", 0, type="TextItem", text="hello"),
                          item("right", 10, revision="r2", type="VideoItem", source_path="C:/b.mp4"),
                          item("other", 20, layer=2, type="TextItem", text="keep"))

    def check(self, after=None, before=None, **kw):
        return verify(self.before if before is None else before,
                      self.after if after is None else after,
                      start_frame=10, end_frame=20, layers=[1],
                      expected_revisions=self.revisions, **kw)

    def test_exact_ripple_passes(self):
        result = self.check()
        self.assertTrue(result["passed"])
        self.assertFalse(result["executed"])
        self.assertFalse(result["atomicity_verified"])
        self.assertEqual((result["planned_delete_count"], result["planned_move_count"]), (1, 1))

    def test_deleted_item_still_present(self):
        after = snap(*self.after["items"], self.before["items"][1])
        self.assertIn("RIPPLE_ITEM_NOT_DELETED",
                      {i["code"] for i in self.check(after=after)["issues"]})

    def test_moved_item_wrong_frame(self):
        after = snap(self.after["items"][0], item("right", 20, type="VideoItem", source_path="C:/b.mp4"),
                     self.after["items"][2])
        self.assertIn("RIPPLE_FIELD_MISMATCH",
                      {i["code"] for i in self.check(after=after)["issues"]})

    def test_unaffected_item_moved(self):
        after = snap(item("left", 1, type="TextItem", text="hello"), *self.after["items"][1:])
        self.assertFalse(self.check(after=after)["passed"])

    def test_unselected_layer_moved(self):
        after = snap(*self.after["items"][:2], item("other", 10, layer=2, type="TextItem", text="keep"))
        self.assertFalse(self.check(after=after)["passed"])

    def test_content_changed_even_when_geometry_matches(self):
        after = snap(*self.after["items"][:2], item("other", 20, layer=2, type="TextItem", text="changed"))
        self.assertIn("RIPPLE_CONTENT_CHANGED",
                      {i["code"] for i in self.check(after=after)["issues"]})

    def test_missing_item(self):
        after = snap(self.after["items"][0], self.after["items"][2])
        self.assertIn("RIPPLE_UNEXPECTED_DELETION",
                      {i["code"] for i in self.check(after=after)["issues"]})

    def test_extra_item(self):
        after = snap(*self.after["items"], item("new", 50))
        self.assertIn("RIPPLE_UNEXPECTED_ADDITION",
                      {i["code"] for i in self.check(after=after)["issues"]})

    def test_duplicate_ids_rejected(self):
        after = snap(*self.after["items"], self.after["items"][1])
        self.assertEqual(self.check(after=after)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_partial_inventory_rejected(self):
        after = {"items": list(self.after["items"]), "count": 4}
        self.assertEqual(self.check(after=after)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_missing_count_rejected(self):
        after = {"items": list(self.after["items"])}
        self.assertEqual(self.check(after=after)["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_before_invalid_rejected(self):
        before = {"items": self.before["items"], "count": True}
        self.assertEqual(self.check(before=before)["error_code"], "RIPPLE_BEFORE_SNAPSHOT_INVALID")

    def test_stale_revision_preflight_fails(self):
        result = verify(self.before, self.after, start_frame=10, end_frame=20,
                        layers=[1], expected_revisions={"cut": "stale", "right": "r1"})
        self.assertEqual(result["error_code"], "RIPPLE_PREFLIGHT_FAILED")
        self.assertFalse(result["passed"])

    def test_crossing_boundary_preflight_fails(self):
        before = snap(item("cross", 5, 20))
        result = verify(before, before, start_frame=10, end_frame=20,
                        expected_revisions={"cross": "r1"})
        self.assertEqual(result["error_code"], "RIPPLE_PREFLIGHT_FAILED")

    def test_snapshot_input_unchanged(self):
        original = repr(self.before), repr(self.after)
        self.check()
        self.assertEqual(original, (repr(self.before), repr(self.after)))


if __name__ == "__main__":
    unittest.main()
