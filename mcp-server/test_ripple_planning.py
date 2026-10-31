"""Regression coverage for the read-only ripple-delete preflight (#61)."""
import unittest
from ripple_planning import plan_ripple_delete as plan


def item(name, frame, length=10, layer=1, revision="r1"):
    return dict(item_id=name, frame=frame, length=length, layer=layer, revision=revision)


class RipplePlanningTests(unittest.TestCase):
    def test_delete_and_shift_without_writes(self):
        before = [item("left", 0), item("cut", 10), item("right", 20), item("other", 20, layer=2)]
        original = [dict(i) for i in before]
        result = plan({"items": before}, start_frame=10, end_frame=20, layers=[1],
                      expected_revisions={"cut": "r1", "right": "r1"})
        self.assertTrue(result["success"])
        self.assertFalse(result["executable"])
        self.assertEqual((result["delete_count"], result["move_count"]), (1, 1))
        self.assertEqual(result["operations"], [
            {"action": "delete", "item_id": "cut", "layer": 1, "frame": 10,
             "length": 10, "expected_revision": "r1"},
            {"action": "move", "item_id": "right", "layer": 1,
             "from_frame": 20, "to_frame": 10, "length": 10, "expected_revision": "r1"}])
        self.assertEqual(before, original)

    def test_cut_middle_requires_split(self):
        result = plan({"items": [item("long", 0, 30)]}, start_frame=10,
                      end_frame=20, expected_revisions={"long": "r1"})
        self.assertEqual(result["error_code"], "RIPPLE_SPLIT_REQUIRED")
        self.assertEqual(result["conflicts"][0]["item_id"], "long")

    def test_cross_left_boundary(self):
        self.assertEqual(plan({"items": [item("a", 5, 10)]}, start_frame=10,
                              end_frame=20)["error_code"], "RIPPLE_SPLIT_REQUIRED")

    def test_cross_right_boundary(self):
        self.assertEqual(plan({"items": [item("a", 15, 10)]}, start_frame=10,
                              end_frame=20)["error_code"], "RIPPLE_SPLIT_REQUIRED")

    def test_boundary_touching_not_split(self):
        result = plan({"items": [item("a", 0, 10), item("b", 20, 10)]},
                      start_frame=10, end_frame=20, expected_revisions={"b": "r1"})
        self.assertEqual(result["move_count"], 1)
        self.assertEqual(result["operations"][0]["to_frame"], 10)

    def test_missing_expected_revision(self):
        result = plan({"items": [item("a", 20)]}, start_frame=10, end_frame=20)
        self.assertEqual(result["error_code"], "RIPPLE_CONFLICT")

    def test_stale_revision(self):
        result = plan({"items": [item("a", 20, revision="r2")]}, start_frame=10,
                      end_frame=20, expected_revisions={"a": "r1"})
        self.assertEqual(result["error_code"], "RIPPLE_CONFLICT")

    def test_duplicate_id_across_layers(self):
        result = plan({"items": [item("a", 0), item("a", 20, layer=2)]},
                      start_frame=10, end_frame=20)
        self.assertEqual(result["error_code"], "RIPPLE_ID_AMBIGUOUS")

    def test_invalid_snapshot_fails_closed(self):
        self.assertEqual(plan({"success": False, "items": []}, start_frame=10,
                              end_frame=20)["error_code"], "RIPPLE_SNAPSHOT_UNAVAILABLE")
        self.assertEqual(plan({"items": [item("a", -1)]}, start_frame=10,
                              end_frame=20)["error_code"], "RIPPLE_SNAPSHOT_INVALID")

    def test_untargeted_crossing_item_ignored(self):
        result = plan({"items": [item("a", 0, 30, layer=2)]}, start_frame=10,
                      end_frame=20, layers=[1])
        self.assertTrue(result["success"])
        self.assertEqual(result["operations"], [])

    def test_invalid_ranges_and_layers(self):
        for kwargs in ({"start_frame": 10, "end_frame": 10},
                       {"start_frame": True, "end_frame": 20},
                       {"start_frame": 10, "end_frame": 20, "layers": [1, 1]},
                       {"start_frame": 10, "end_frame": 20, "layers": [True]}):
            with self.subTest(kwargs=kwargs), self.assertRaises(ValueError):
                plan({"items": []}, **kwargs)

    def test_canonical_order_independent_of_snapshot_order(self):
        items = [item("z", 20, layer=2), item("a", 10, layer=1), item("m", 30, layer=1)]
        expected = {i["item_id"]: "r1" for i in items}
        a = plan({"items": items}, start_frame=10, end_frame=20,
                 expected_revisions=expected)
        b = plan({"items": list(reversed(items))}, start_frame=10, end_frame=20,
                 expected_revisions=expected)
        self.assertEqual(a["operations"], b["operations"])

    def test_noop_requires_no_revisions(self):
        result = plan({"items": [item("left", 0)]}, start_frame=20, end_frame=30)
        self.assertTrue(result["success"])
        self.assertEqual(result["operations"], [])


if __name__ == "__main__":
    unittest.main()
