"""Regression tests for scene summary integrity and cross-scene trust boundaries."""
import copy
import unittest

import scenes
from editing import MAX_FRAME
from scene_media import MAX_TOTAL_SCENE_ITEMS, with_coverage


def snapshot():
    return scenes.index(
        [{"item_id": "video", "frame": 0, "length": 30,
          "layer": 1, "type": "VideoItem"},
         {"item_id": "audio", "frame": 10, "length": 10,
          "layer": 2, "type": "AudioItem"},
         {"item_id": "unassigned", "frame": 80, "length": 10,
          "layer": 3, "type": "TextItem"}],
        [{"id": "intro", "start_frame": 0, "end_frame": 20},
         {"id": "outro", "start_frame": 20, "end_frame": 40}])


class SceneEvidenceSummaryTests(unittest.TestCase):
    def test_accepts_real_scene_index_with_crossing_and_unassigned_items(self):
        index = snapshot()
        result = with_coverage(index)
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)
        self.assertEqual(result["scenes"][1]["visual_covered_frames"], 10)
        self.assertEqual(result["scenes"][0]["audio_covered_frames"], 10)
        self.assertEqual(result["scenes"][1]["audio_covered_frames"], 0)
        self.assertEqual(result["unassigned_item_ids"], ["unassigned"])

    def test_rejects_unsuccessful_index_before_calculating_coverage(self):
        for value in (False, None, 0, "true", {}, []):
            with self.subTest(value=value):
                index = snapshot()
                index["success"] = value
                with self.assertRaisesRegex(ValueError, "report success"):
                    with_coverage(index)

    def test_rejects_forged_scene_covered_frame_total(self):
        for value in (19, True, "20", None):
            with self.subTest(value=value):
                index = snapshot()
                index["scenes"][0]["covered_frames"] = value
                with self.assertRaisesRegex(ValueError, "covered_frames"):
                    with_coverage(index)

    def test_rejects_forged_scene_uncovered_frame_total(self):
        for value in (1, False, "0", None):
            with self.subTest(value=value):
                index = snapshot()
                index["scenes"][0]["uncovered_frames"] = value
                with self.assertRaisesRegex(ValueError, "uncovered_frames"):
                    with_coverage(index)

    def test_rejects_forged_scene_gap_boundaries(self):
        index = snapshot()
        index["scenes"][1]["uncovered_ranges"] = [
            {"start_frame": 30, "end_frame": 39}]
        with self.assertRaisesRegex(ValueError, "uncovered_ranges"):
            with_coverage(index)

    def test_rejects_forged_layer_occupancy_summary(self):
        for field, bad in (
            ("covered_frames", 0), ("uncovered_frames", 20),
            ("item_ids", ["audio"]), ("uncovered_ranges", []),
            ("layer", 9)):
            with self.subTest(field=field):
                index = snapshot()
                index["scenes"][0]["layer_occupancy"][0][field] = bad
                with self.assertRaisesRegex(ValueError, "layer_occupancy"):
                    with_coverage(index)

    def test_rejects_missing_or_extra_layer_occupancy_rows(self):
        for change in ("drop", "extra"):
            with self.subTest(change=change):
                index = snapshot()
                rows = index["scenes"][0]["layer_occupancy"]
                if change == "drop":
                    rows.pop()
                else:
                    rows.append(copy.deepcopy(rows[0]))
                with self.assertRaisesRegex(ValueError, "layer_occupancy"):
                    with_coverage(index)

    def test_rejects_forged_previous_continuation(self):
        index = snapshot()
        index["scenes"][1]["continues_from_previous_item_ids"] = []
        with self.assertRaisesRegex(ValueError, "continues_from_previous"):
            with_coverage(index)

    def test_rejects_forged_next_continuation(self):
        index = snapshot()
        index["scenes"][0]["continues_into_next_item_ids"] = ["audio"]
        with self.assertRaisesRegex(ValueError, "continues_into_next"):
            with_coverage(index)

    def test_rejects_continuation_without_source_identity(self):
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20, "items": [
                {"type": "VideoItem", "visible_start_frame": 0,
                 "visible_end_frame": 20}],
             "continues_into_next_item_ids": ["v"]},
            {"start_frame": 20, "end_frame": 30, "items": []}]}
        with self.assertRaisesRegex(ValueError, "continuation requires"):
            with_coverage(index)

    def test_rejects_out_of_range_layer(self):
        for bad in (-1, MAX_FRAME + 1):
            with self.subTest(bad=bad):
                index = {"scenes": [
                    {"start_frame": 0, "end_frame": 20,
                     "items": [{"type": "Unknown", "layer": bad}]}]}
                with self.assertRaisesRegex(ValueError, "item layer"):
                    with_coverage(index)

    def test_unknown_type_cannot_claim_non_intersecting_scene_membership(self):
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20, "items": [
                {"type": "Unknown", "item_id": "x",
                 "frame": 30, "end_frame": 40,
                 "visible_start_frame": 30, "visible_end_frame": 20}]}]}
        with self.assertRaisesRegex(ValueError, "does not intersect"):
            with_coverage(index)

    def test_rejects_partial_identity_change_across_scenes(self):
        first = {"item_id": "same", "type": "Unknown"}
        second = {"item_id": "same", "type": "Unknown",
                  "frame": 0, "end_frame": 30,
                  "visible_start_frame": 20, "visible_end_frame": 30}
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20, "items": [first]},
            {"start_frame": 20, "end_frame": 30, "items": [second]}]}
        with self.assertRaisesRegex(ValueError, "conflicting source evidence"):
            with_coverage(index)

    def test_rejects_changed_type_without_source_spans(self):
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 20,
             "items": [{"item_id": "same", "type": "Unknown"}]},
            {"start_frame": 20, "end_frame": 30,
             "items": [{"item_id": "same", "type": "VideoItem",
                        "visible_start_frame": 20, "visible_end_frame": 30}]}]}
        with self.assertRaisesRegex(ValueError, "conflicting source evidence"):
            with_coverage(index)

    def test_rejects_oversized_total_even_with_unassigned_items(self):
        index = {"scenes": [{"start_frame": 0, "end_frame": 20, "items": []}],
                 "item_count": MAX_TOTAL_SCENE_ITEMS + 1,
                 "unassigned_item_ids": []}
        with self.assertRaisesRegex(ValueError, "index item count"):
            with_coverage(index)

    def test_failed_summary_validation_preserves_original_snapshot(self):
        index = snapshot()
        index["scenes"][1]["covered_frames"] = 999
        before = copy.deepcopy(index)
        with self.assertRaises(ValueError):
            with_coverage(index)
        self.assertEqual(index, before)


if __name__ == "__main__":
    unittest.main()
