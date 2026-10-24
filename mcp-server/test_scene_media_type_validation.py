"""Scene occupancy must not mistake unknown type names for known media items."""
import unittest

import scene_media


class SceneMediaTypeValidationTests(unittest.TestCase):
    def coverage(self, kind):
        scene = {"start_frame": 0, "end_frame": 20,
                 "items": [{"type": kind, "visible_start_frame": 0,
                            "visible_end_frame": 20}]}
        return scene_media.with_coverage({"scenes": [scene]})["scenes"][0]

    def test_unknown_prefixed_video_type_does_not_hide_visual_gap(self):
        scene = self.coverage("NotVideoItem")
        self.assertEqual(scene["visual_covered_frames"], 0)
        self.assertEqual(scene["visual_uncovered_ranges"],
                         [{"start_frame": 0, "end_frame": 20}])

    def test_unknown_prefixed_voice_type_does_not_hide_audio_gap(self):
        scene = self.coverage("BrokenVoiceItem")
        self.assertEqual(scene["audio_covered_frames"], 0)

    def test_known_namespaced_type_is_recognized(self):
        scene = self.coverage("YukkuriMovieMaker.Project.Items.VideoItem")
        self.assertEqual(scene["visual_covered_frames"], 20)

    def test_known_nested_type_is_recognized(self):
        scene = self.coverage("YukkuriMovieMaker.Items+VoiceItem")
        self.assertEqual(scene["audio_covered_frames"], 20)

    def test_case_insensitive_known_type(self):
        scene = self.coverage("VIDEOITEM")
        self.assertEqual(scene["visual_covered_frames"], 20)

    def test_missing_type_remains_uncovered(self):
        scene = self.coverage(None)
        self.assertEqual(scene["visual_covered_frames"], 0)
        self.assertEqual(scene["audio_covered_frames"], 0)


if __name__ == "__main__":
    unittest.main()
