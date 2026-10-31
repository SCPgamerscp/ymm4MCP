"""Regression tests for canonical scene evidence and Unicode identity safety."""
import copy
import unittest

import scenes
from scene_media import with_coverage


RANGES = [
    {"id": "intro", "start_frame": 0, "end_frame": 20},
    {"id": "body", "start_frame": 20, "end_frame": 40},
]


def canonical():
    return scenes.index(
        [{"item_id": "video", "frame": 0, "length": 30,
          "layer": 1, "type": "VideoItem"}], RANGES)


def legacy(items):
    return {"scenes": [
        {"start_frame": 0, "end_frame": 20, "items": items},
    ]}


class CanonicalSceneTrustTests(unittest.TestCase):
    def test_complete_canonical_snapshot_remains_valid_and_idempotent(self):
        index = canonical()
        self.assertIs(with_coverage(index), index)
        self.assertIs(with_coverage(index), index)
        self.assertEqual(index["scenes"][1]["visual_covered_frames"], 10)

    def test_canonical_scene_cannot_omit_evidence_fields(self):
        for key in ("id", "item_count", "continues_from_previous_item_ids",
                    "continues_into_next_item_ids", "covered_frames",
                    "uncovered_frames", "uncovered_ranges", "layer_occupancy"):
            with self.subTest(key=key):
                index = canonical()
                del index["scenes"][0][key]
                before = copy.deepcopy(index)
                with self.assertRaisesRegex(ValueError, "canonical scene snapshot"):
                    with_coverage(index)
                self.assertEqual(index, before)

    def test_canonical_item_cannot_omit_evidence_fields(self):
        for key in ("item_id", "frame", "end_frame", "layer", "type",
                    "visible_start_frame", "visible_end_frame",
                    "relative_start_frame", "relative_end_frame",
                    "crosses_boundary"):
            with self.subTest(key=key):
                index = canonical()
                del index["scenes"][0]["items"][0][key]
                with self.assertRaisesRegex(ValueError, "canonical scene item"):
                    with_coverage(index)

    def test_legacy_minimal_index_still_supported(self):
        result = with_coverage(legacy([
            {"type": "VideoItem", "visible_start_frame": 0,
             "visible_end_frame": 20},
        ]))
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_rejects_error_metadata_even_with_success_true(self):
        for error in ("error", "error_code"):
            with self.subTest(error=error):
                index = canonical()
                index[error] = "ITEMS_UNAVAILABLE"
                with self.assertRaisesRegex(ValueError, "host error"):
                    with_coverage(index)

    def test_rejects_ambiguous_repeated_legacy_id_without_source_frames(self):
        index = {"scenes": [
            {"start_frame": 0, "end_frame": 10, "items": [
                {"item_id": "same", "type": "VideoItem",
                 "visible_start_frame": 0, "visible_end_frame": 10}]},
            {"start_frame": 10, "end_frame": 20, "items": [
                {"item_id": "same", "type": "VideoItem",
                 "visible_start_frame": 10, "visible_end_frame": 20}]},
        ]}
        with self.assertRaisesRegex(ValueError, "requires source frames"):
            with_coverage(index)

    def test_single_legacy_id_without_source_frames_is_allowed(self):
        result = with_coverage(legacy([
            {"item_id": "legacy", "type": "VideoItem",
             "visible_start_frame": 0, "visible_end_frame": 20},
        ]))
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_rejects_padded_ids_at_index_construction(self):
        for key, value in (("item_id", " video"), ("item_id", "video "),
                           ("scene_id", " intro"), ("scene_id", "intro ")):
            with self.subTest(key=key, value=value):
                items = [{"item_id": "video", "frame": 0, "length": 1}]
                ranges = copy.deepcopy(RANGES)
                if key == "item_id":
                    items[0]["item_id"] = value
                else:
                    ranges[0]["id"] = value
                with self.assertRaises(ValueError):
                    scenes.index(items, ranges)

    def test_rejects_padded_ids_in_external_scene_evidence(self):
        for item_id in (" video", "video "):
            with self.subTest(item_id=item_id):
                index = legacy([
                    {"item_id": item_id, "type": "VideoItem",
                     "visible_start_frame": 0, "visible_end_frame": 20},
                ])
                with self.assertRaisesRegex(ValueError, "item IDs"):
                    with_coverage(index)

    def test_rejects_invisible_format_and_surrogate_ids(self):
        for identity in ("a\u200bb", "a\u2060b", "a\ufeffb", "a\ud800b"):
            with self.subTest(identity=repr(identity)):
                with self.assertRaises(ValueError):
                    scenes.index([{"item_id": identity, "frame": 0,
                                   "length": 1}], RANGES)
                with self.assertRaisesRegex(ValueError, "item IDs"):
                    with_coverage(legacy([
                        {"item_id": identity, "type": "VideoItem",
                         "visible_start_frame": 0, "visible_end_frame": 20},
                    ]))

    def test_rejects_invisible_format_and_surrogate_item_types(self):
        for kind in ("Video\u200bItem", "Video\u2060Item",
                     "Video\ufeffItem", "Video\ud800Item"):
            with self.subTest(kind=repr(kind)):
                with self.assertRaisesRegex(ValueError, "item.type"):
                    scenes.index([{"item_id": "v", "frame": 0, "length": 1,
                                   "type": kind}], RANGES)
                with self.assertRaisesRegex(ValueError, "item type"):
                    with_coverage(legacy([
                        {"type": kind, "visible_start_frame": 0,
                         "visible_end_frame": 20},
                    ]))

    def test_rejects_non_nfc_identifiers_but_accepts_composed_unicode(self):
        decomposed = "e\u0301"
        with self.assertRaises(ValueError):
            scenes.index([{"item_id": decomposed, "frame": 0,
                           "length": 1}], RANGES)
        with self.assertRaisesRegex(ValueError, "item IDs"):
            with_coverage(legacy([
                {"item_id": decomposed, "type": "VideoItem",
                 "visible_start_frame": 0, "visible_end_frame": 20},
            ]))
        index = scenes.index([{"item_id": "\u00e9", "frame": 0,
                               "length": 1, "type": "VideoItem"}], RANGES)
        self.assertEqual(with_coverage(index)["scenes"][0]["visual_covered_frames"], 1)


if __name__ == "__main__":
    unittest.main()
