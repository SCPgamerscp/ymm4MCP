"""Regression: invisible Unicode controls must not spoof scene/item identities."""
import unittest

from scene_media import with_coverage


class SceneIdentifierFormatControlTests(unittest.TestCase):
    def test_invisible_controls_rejected_in_scene_id(self):
        for control in ("\u200c", "\u200d", "\u200e", "\u2060"):
            with self.subTest(control=repr(control)):
                index = {"scenes": [{
                    "id": "scene" + control + "1",
                    "start_frame": 0, "end_frame": 10, "items": [],
                }]}
                with self.assertRaisesRegex(ValueError, "scene IDs"):
                    with_coverage(index)

    def test_invisible_controls_rejected_in_item_id(self):
        for control in ("\u200c", "\u200d", "\u200e", "\u2060"):
            with self.subTest(control=repr(control)):
                index = {"scenes": [{
                    "start_frame": 0, "end_frame": 10,
                    "items": [{"item_id": "clip" + control + "1",
                               "type": "VideoItem",
                               "visible_start_frame": 0,
                               "visible_end_frame": 10}],
                }]}
                with self.assertRaisesRegex(ValueError, "scene item IDs"):
                    with_coverage(index)

    def test_normal_unicode_identifiers_remain_valid(self):
        index = {"scenes": [{
            "id": "場面一", "start_frame": 0, "end_frame": 10,
            "items": [{"item_id": "映像一", "type": "VideoItem",
                       "visible_start_frame": 0, "visible_end_frame": 10}],
        }]}
        self.assertEqual(with_coverage(index)["scenes"][0]["visual_covered_frames"], 10)


if __name__ == "__main__":
    unittest.main()
