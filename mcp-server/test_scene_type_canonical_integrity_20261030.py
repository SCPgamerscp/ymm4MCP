"""Regression tests for complete scene snapshots and Unicode-safe type identity."""
import copy
import unittest

import scenes
from scene_media import with_coverage


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 20}]


def source(item_id="video", kind="VideoItem"):
    return scenes.index(
        [{"item_id": item_id, "frame": 0, "length": 20,
          "layer": 1, "type": kind}], RANGES)


def legacy(item_id="video", kind="VideoItem"):
    return {"scenes": [{"start_frame": 0, "end_frame": 20,
                        "items": [{"item_id": item_id, "type": kind,
                                   "visible_start_frame": 0,
                                   "visible_end_frame": 20}]}]}


class SceneTypeCanonicalIntegrityTests(unittest.TestCase):
    def test_success_snapshot_missing_item_count_fails_closed(self):
        data = source()
        del data["item_count"]
        before = copy.deepcopy(data)
        with self.assertRaisesRegex(ValueError, "top-level evidence"):
            with_coverage(data)
        self.assertEqual(data, before)

    def test_success_snapshot_missing_unassigned_ids_fails_closed(self):
        data = source()
        del data["unassigned_item_ids"]
        with self.assertRaisesRegex(ValueError, "top-level evidence"):
            with_coverage(data)

    def test_success_snapshot_cannot_downgrade_to_legacy(self):
        data = source()
        del data["item_count"]
        del data["unassigned_item_ids"]
        del data["scenes"][0]["layer_occupancy"]
        with self.assertRaisesRegex(ValueError, "top-level evidence"):
            with_coverage(data)

    def test_minimal_legacy_snapshot_remains_accepted(self):
        data = legacy()
        self.assertEqual(with_coverage(data)["scenes"][0]["visual_covered_frames"], 20)

    def test_valid_complete_snapshot_is_idempotent(self):
        data = source()
        with_coverage(data)
        before = copy.deepcopy(data)
        self.assertIs(with_coverage(data), data)
        self.assertEqual(data, before)

    def test_source_rejects_padded_known_types(self):
        for kind in (" VideoItem", "VideoItem ", "VideoItem\u00a0"):
            with self.subTest(kind=repr(kind)), self.assertRaisesRegex(ValueError, "item.type"):
                source(kind=kind)

    def test_coverage_rejects_padded_known_types(self):
        for kind in (" VideoItem", "VideoItem ", "VideoItem\u00a0"):
            with self.subTest(kind=repr(kind)), self.assertRaisesRegex(ValueError, "item type"):
                with_coverage(legacy(kind=kind))

    def test_source_rejects_invisible_joiners_in_type(self):
        for kind in ("Video\u200cItem", "Video\u200dItem"):
            with self.subTest(kind=repr(kind)), self.assertRaisesRegex(ValueError, "item.type"):
                source(kind=kind)

    def test_coverage_rejects_invisible_joiners_in_type(self):
        for kind in ("Video\u200cItem", "Video\u200dItem"):
            with self.subTest(kind=repr(kind)), self.assertRaisesRegex(ValueError, "item type"):
                with_coverage(legacy(kind=kind))

    def test_source_rejects_non_nfc_type(self):
        with self.assertRaisesRegex(ValueError, "item.type"):
            source(kind="Vide\u006f\u0301Item")

    def test_coverage_rejects_non_nfc_type(self):
        with self.assertRaisesRegex(ValueError, "item type"):
            with_coverage(legacy(kind="Vide\u006f\u0301Item"))

    def test_source_rejects_unicode_line_separator_in_identity(self):
        for item_id in ("video\u2028one", "video\u2029one"):
            with self.subTest(item_id=repr(item_id)), self.assertRaisesRegex(ValueError, "item_id"):
                source(item_id=item_id)

    def test_coverage_rejects_unicode_line_separator_in_identity(self):
        for item_id in ("video\u2028one", "video\u2029one"):
            with self.subTest(item_id=repr(item_id)), self.assertRaisesRegex(ValueError, "item IDs"):
                with_coverage(legacy(item_id=item_id))

    def test_noncharacters_rejected_in_source_ids_and_types(self):
        for char in ("\ufdd0", "\ufffe", "\U0010ffff"):
            with self.subTest(char=repr(char)):
                with self.assertRaisesRegex(ValueError, "item_id"):
                    source(item_id="v" + char)
                with self.assertRaisesRegex(ValueError, "item.type"):
                    source(kind="VideoItem" + char)

    def test_noncharacters_rejected_in_coverage_ids_and_types(self):
        for char in ("\ufdd0", "\ufffe", "\U0010ffff"):
            with self.subTest(char=repr(char)):
                with self.assertRaisesRegex(ValueError, "item IDs"):
                    with_coverage(legacy(item_id="v" + char))
                with self.assertRaisesRegex(ValueError, "item type"):
                    with_coverage(legacy(kind="VideoItem" + char))

    def test_emoji_joiner_in_identifier_still_allowed(self):
        data = source(item_id="emoji-\U0001f469\u200d\U0001f4bb")
        self.assertEqual(with_coverage(data)["scenes"][0]["visual_covered_frames"], 20)

    def test_namespaced_known_type_still_provides_visual_coverage(self):
        data = source(kind="YMM4.Items+VideoItem")
        self.assertEqual(with_coverage(data)["scenes"][0]["visual_covered_frames"], 20)


if __name__ == "__main__":
    unittest.main()
