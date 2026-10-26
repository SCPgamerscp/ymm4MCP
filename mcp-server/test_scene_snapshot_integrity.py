"""Fail closed on ambiguous IDs and malformed native item types in scene QA."""
import unittest

import scenes


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 10}]
ITEM = {"item_id": "native:1", "frame": 0, "length": 10, "layer": 1,
        "type": "ImageItem"}


class SceneSnapshotIntegrityTests(unittest.TestCase):
    def test_rejects_bidirectional_controls_and_c1_ids(self):
        invalid = ("\u0085", "\u009f", "\u061c", "\u200e", "\u200f",
                   "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",
                   "\u2066", "\u2067", "\u2068", "\u2069")
        for mark in invalid:
            with self.subTest(mark=hex(ord(mark))):
                with self.assertRaisesRegex(ValueError, "scene ids"):
                    scenes.index([], [{"id": "intro" + mark, "start_frame": 0,
                                       "end_frame": 10}])
                with self.assertRaisesRegex(ValueError, "item_id"):
                    scenes.index([{**ITEM, "item_id": "native" + mark + ":1"}], RANGES)

    def test_rejects_unbounded_native_ids(self):
        with self.assertRaisesRegex(ValueError, "item_id"):
            scenes.index([{**ITEM, "item_id": "n" * 257}], RANGES)
        result = scenes.index([{**ITEM, "item_id": "n" * 256}], RANGES)
        self.assertEqual(result["item_count"], 1)

    def test_rejects_malformed_explicit_native_types(self):
        for bad in (None, True, 1, [], {}, "", "   ", "A" * 129):
            with self.subTest(type=repr(bad)), self.assertRaisesRegex(ValueError, "item.type"):
                scenes.index([{**ITEM, "type": bad}], RANGES)

    def test_preserves_missing_types_and_readable_unicode(self):
        item = {key: value for key, value in ITEM.items() if key != "type"}
        result = scenes.index([{**item, "item_id": "native:立ち絵"}],
                              [{"id": "家族👨\u200d👩\u200d👧", "start_frame": 0, "end_frame": 10}])
        self.assertEqual(result["scenes"][0]["items"][0]["item_id"], "native:立ち絵")
        self.assertEqual(result["scenes"][0]["items"][0]["type"], None)
        self.assertEqual(scenes.index([ITEM], RANGES)["scenes"][0]["item_count"], 1)


if __name__ == "__main__":
    unittest.main()
