"""Regression tests for scene snapshot ordering and derived metadata integrity."""
import copy
import unittest

from scene_media import MAX_SCENE_ITEMS, MAX_TOTAL_SCENE_ITEMS, with_coverage


def scene(start=0, end=20, items=None, **extra):
    return {"start_frame": start, "end_frame": end,
            "items": [] if items is None else items, **extra}


def video(identity="v", frame=0, end=20, **extra):
    return {"item_id": identity, "type": "VideoItem",
            "frame": frame, "end_frame": end,
            "visible_start_frame": max(frame, 0),
            "visible_end_frame": min(end, 20), **extra}


class SceneCoverageSnapshotIntegrityTests(unittest.TestCase):
    def test_rejects_overlapping_scenes(self):
        index = {"scenes": [scene(0, 20), scene(19, 30)]}
        with self.assertRaisesRegex(ValueError, "non-overlapping"):
            with_coverage(index)

    def test_rejects_out_of_order_scenes(self):
        index = {"scenes": [scene(20, 30), scene(0, 10)]}
        with self.assertRaisesRegex(ValueError, "ordered"):
            with_coverage(index)

    def test_accepts_adjacent_and_separated_scenes(self):
        index = {"scenes": [scene(0, 10), scene(10, 20), scene(30, 40)]}
        result = with_coverage(index)
        self.assertEqual([s["visual_covered_frames"] for s in result["scenes"]],
                         [0, 0, 0])

    def test_aggregate_scene_item_budget_rejects_excess(self):
        self.assertEqual(MAX_TOTAL_SCENE_ITEMS, 2 * MAX_SCENE_ITEMS)
        index = {"scenes": [
            scene(0, 10, [{}] * MAX_SCENE_ITEMS),
            scene(10, 20, [{}] * MAX_SCENE_ITEMS),
            scene(20, 30, [{}])]}
        with self.assertRaisesRegex(ValueError, "aggregate item budget"):
            with_coverage(index)

    def test_relative_frames_accept_consistent_evidence(self):
        item = video(relative_start_frame=0, relative_end_frame=20)
        result = with_coverage({"scenes": [scene(items=[item])]})
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_relative_frames_reject_missing_partner(self):
        for field in ("relative_start_frame", "relative_end_frame"):
            with self.subTest(field=field), self.assertRaisesRegex(
                    ValueError, "relative item range"):
                with_coverage({"scenes": [scene(items=[video(**{field: 0})])]})

    def test_relative_frames_reject_wrong_offsets_and_types(self):
        for extra in ({"relative_start_frame": 1, "relative_end_frame": 20},
                      {"relative_start_frame": 0, "relative_end_frame": 19},
                      {"relative_start_frame": True, "relative_end_frame": 20},
                      {"relative_start_frame": 0, "relative_end_frame": "20"}):
            with self.subTest(extra=extra), self.assertRaisesRegex(
                    ValueError, "relative item range"):
                with_coverage({"scenes": [scene(items=[video(**extra)])]})

    def test_relative_frames_support_scene_offsets(self):
        item = {"type": "AudioItem", "visible_start_frame": 12,
                "visible_end_frame": 19, "relative_start_frame": 2,
                "relative_end_frame": 9}
        result = with_coverage({"scenes": [scene(10, 20, [item])]})
        self.assertEqual(result["scenes"][0]["audio_covered_frames"], 7)

    def test_boundary_flag_accepts_true_and_false(self):
        crossing = video(frame=0, end=30, visible_end_frame=20,
                         crosses_boundary=True)
        internal = video(identity="b", crosses_boundary=False)
        result = with_coverage({"scenes": [scene(items=[crossing, internal])]})
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_boundary_flag_rejects_wrong_value(self):
        for flag in (True, None, 0, "false"):
            with self.subTest(flag=flag), self.assertRaisesRegex(
                    ValueError, "boundary flag"):
                with_coverage({"scenes": [scene(items=[video(crosses_boundary=flag)])]})

    def test_boundary_flag_requires_source_evidence(self):
        item = {"type": "VideoItem", "visible_start_frame": 0,
                "visible_end_frame": 20, "crosses_boundary": False}
        with self.assertRaisesRegex(ValueError, "boundary flag"):
            with_coverage({"scenes": [scene(items=[item])]})

    def test_index_count_includes_unassigned_and_deduplicates_continuations(self):
        first = {"item_id": "long", "type": "VideoItem", "frame": 0,
                 "end_frame": 30, "visible_start_frame": 0,
                 "visible_end_frame": 20}
        second = {**first, "visible_start_frame": 20,
                  "visible_end_frame": 30}
        index = {"scenes": [scene(0, 20, [first]), scene(20, 30, [second])],
                 "item_count": 2, "unassigned_item_ids": ["offscreen"]}
        result = with_coverage(index)
        self.assertEqual(result["scenes"][1]["visual_covered_frames"], 10)

    def test_index_count_rejects_inconsistent_total(self):
        for total in (0, 2, True, 1.0, "1", -1):
            with self.subTest(total=total), self.assertRaisesRegex(
                    ValueError, "index item count"):
                with_coverage({"scenes": [scene(items=[video()])],
                               "item_count": total, "unassigned_item_ids": []})

    def test_index_rejects_unassigned_duplicates_and_overlap(self):
        for unassigned in (["a", "a"], ["v"], [" "], [None],
                           ["x\n"], [True], ["x" * 257]):
            with self.subTest(unassigned=unassigned), self.assertRaisesRegex(
                    ValueError, "unassigned identities"):
                with_coverage({"scenes": [scene(items=[video()])],
                               "item_count": 1 + len(unassigned),
                               "unassigned_item_ids": unassigned})

    def test_index_requires_complete_metadata_pair(self):
        for extra in ({"item_count": 1}, {"unassigned_item_ids": []},
                      {"item_count": 1, "unassigned_item_ids": None}):
            with self.subTest(extra=extra), self.assertRaisesRegex(
                    ValueError, "index item count"):
                with_coverage({"scenes": [scene(items=[video()])], **extra})

    def test_invalid_snapshot_does_not_mutate_earlier_coverage(self):
        index = {"scenes": [scene(items=[video()])],
                 "item_count": 99, "unassigned_item_ids": []}
        before = copy.deepcopy(index)
        with self.assertRaises(ValueError):
            with_coverage(index)
        self.assertEqual(index, before)


if __name__ == "__main__":
    unittest.main()
