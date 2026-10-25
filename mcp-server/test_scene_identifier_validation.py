"""Regression tests for identifiers returned by scene indexing and QA."""

import unittest

import scenes


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 10}]
ITEM = {"item_id": "native:sample", "frame": 0, "length": 10, "layer": 1}


class SceneIdentifierValidationTests(unittest.TestCase):
    def test_rejects_blank_and_control_character_scene_ids(self):
        for bad in ("", "   ", "\n", "intro\n", "\x00intro", "intro\t", "intro\x7f"):
            with self.subTest(scene_id=repr(bad)), self.assertRaisesRegex(ValueError, "scene ids"):
                scenes.index([], [{"id": bad, "start_frame": 0, "end_frame": 10}])

    def test_rejects_control_character_item_ids(self):
        for bad in ("", "   ", "\n", "native:sample\n", "native:\x00sample",
                    "native:sample\r", "native:sample\x7f"):
            with self.subTest(item_id=repr(bad)), self.assertRaisesRegex(ValueError, "item_id"):
                scenes.index([{**ITEM, "item_id": bad}], RANGES)

    def test_preserves_readable_unicode_and_boundary_ids(self):
        result = scenes.index(
            [{**ITEM, "item_id": "native:立ち絵"}],
            [{"id": "序章 🎬", "start_frame": 0, "end_frame": 10}])
        self.assertEqual(result["scenes"][0]["items"][0]["item_id"], "native:立ち絵")
        self.assertEqual(result["scenes"][0]["id"], "序章 🎬")
        self.assertEqual(scenes.index([], [{"id": "a" * 64, "start_frame": 0,
                                            "end_frame": 10}])["scenes"][0]["id"], "a" * 64)


if __name__ == "__main__":
    unittest.main()
