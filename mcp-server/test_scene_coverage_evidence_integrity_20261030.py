"""Regression tests for scene coverage evidence and trusted type identities."""
import copy
import unittest

from editing import MAX_FRAME
from scene_media import MAX_SCENE_ITEMS, with_coverage


def scene(items=None, *, start=0, end=20, **extra):
    return {"start_frame": start, "end_frame": end,
            "items": [] if items is None else items, **extra}


def video(kind="VideoItem"):
    return {"type": kind, "visible_start_frame": 0,
            "visible_end_frame": 20}


class SceneCoverageEvidenceIntegrityTests(unittest.TestCase):
    def test_rejects_malformed_index_containers(self):
        for value in (None, [], {}, {"scenes": None}, {"scenes": {}},
                      {"scenes": []}):
            with self.subTest(value=value), self.assertRaises(ValueError):
                with_coverage(value)

    def test_rejects_too_many_scenes(self):
        with self.assertRaisesRegex(ValueError, "1..200"):
            with_coverage({"scenes": [scene() for _ in range(201)]})

    def test_rejects_invalid_scene_rows(self):
        for invalid in (None, [], "scene", 1):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                with_coverage({"scenes": [invalid]})

    def test_rejects_negative_or_oversized_scene_frames(self):
        for start, end in ((-1, 20), (0, MAX_FRAME + 1)):
            with self.subTest(start=start, end=end), self.assertRaises(ValueError):
                with_coverage({"scenes": [scene(start=start, end=end)]})

    def test_rejects_missing_or_invalid_scene_item_array(self):
        for invalid in (None, {}, "items", 5):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                with_coverage({"scenes": [scene(invalid)]})
        with self.assertRaises(ValueError):
            with_coverage({"scenes": [{"start_frame": 0, "end_frame": 20}]})

    def test_rejects_oversized_scene_item_array(self):
        with self.assertRaisesRegex(ValueError, "bounded array"):
            with_coverage({"scenes": [scene([{}] * (MAX_SCENE_ITEMS + 1))]})

    def test_rejects_non_object_items_even_when_unknown(self):
        for invalid in (None, 2, "item", []):
            with self.subTest(invalid=invalid), self.assertRaisesRegex(
                    ValueError, "scene item must be an object"):
                with_coverage({"scenes": [scene([invalid])]})

    def test_rejects_contradictory_item_count(self):
        for invalid in (0, 2, True, 1.0, "1", None):
            with self.subTest(invalid=invalid), self.assertRaisesRegex(
                    ValueError, "item_count"):
                with_coverage({"scenes": [scene([video()], item_count=invalid)]})

    def test_accepts_consistent_item_count(self):
        result = with_coverage({"scenes": [scene([video()], item_count=1)]})
        self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_rejects_duplicate_or_invalid_scene_ids(self):
        for invalid in ("", " ", "x" * 65, None):
            with self.subTest(invalid=invalid), self.assertRaisesRegex(
                    ValueError, "scene IDs"):
                with_coverage({"scenes": [scene(id=invalid)]})
        with self.assertRaisesRegex(ValueError, "scene IDs"):
            with_coverage({"scenes": [scene(id="intro"), scene(id="intro")]})

    def test_malformed_type_qualifiers_cannot_hide_gaps(self):
        for kind in ("Other..VideoItem", ".VideoItem", "Other+.VideoItem",
                     " VideoItem", "VideoItem ", "Other.\\nVideoItem",
                     "Other.VideoItem+"):
            with self.subTest(kind=kind):
                result = with_coverage({"scenes": [scene([video(kind)])]})
                self.assertEqual(result["scenes"][0]["visual_covered_frames"], 0)

    def test_legitimate_namespaced_and_nested_types_still_cover(self):
        for kind in ("VideoItem", "VIDEOITEM",
                     "YukkuriMovieMaker.Project.Items.VideoItem",
                     "YukkuriMovieMaker.Items+VideoItem"):
            with self.subTest(kind=kind):
                result = with_coverage({"scenes": [scene([video(kind)])]})
                self.assertEqual(result["scenes"][0]["visual_covered_frames"], 20)

    def test_later_invalid_scene_does_not_mutate_earlier_scene(self):
        index = {"scenes": [scene([video()], id="first"),
                            scene([video()], id="first")]}
        original = copy.deepcopy(index)
        with self.assertRaises(ValueError):
            with_coverage(index)
        self.assertEqual(index, original)


if __name__ == "__main__":
    unittest.main()
