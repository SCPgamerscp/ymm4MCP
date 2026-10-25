"""Localized thumbnail motion must not be misreported as a static preview."""
import unittest

import visual_qa


PIXEL_COUNT = 64 * 36


def thumbnail_with_dots(indices):
    pixels = bytearray([128] * (PIXEL_COUNT * 3))
    for pixel in indices:
        pixels[pixel * 3:pixel * 3 + 3] = b"\xff\xff\xff"
    return bytes(pixels)


class LocalizedMotionTests(unittest.TestCase):
    def test_small_moving_detail_breaks_static_run(self):
        base = thumbnail_with_dots([])
        changed = thumbnail_with_dots([0, 1, 64, 65])
        result = visual_qa.inspect(
            [(0, base), (30, base), (60, changed), (90, changed)],
            step_frames=30, min_static_frames=60)
        self.assertNotIn("STATIC_PREVIEW", [issue["code"] for issue in result["issues"]])

    def test_single_noisy_pixel_does_not_break_static_run(self):
        base = thumbnail_with_dots([])
        noise = thumbnail_with_dots([0])
        result = visual_qa.inspect(
            [(0, base), (30, noise), (60, base)],
            step_frames=30, min_static_frames=60)
        static = [issue for issue in result["issues"] if issue["code"] == "STATIC_PREVIEW"]
        self.assertEqual([(issue["start_frame"], issue["end_frame"]) for issue in static],
                         [(0, 60)])

    def test_gradual_fade_still_counts_as_motion(self):
        samples = [(frame, bytes([level] * (PIXEL_COUNT * 3)))
                   for frame, level in ((0, 128), (30, 130), (60, 132), (90, 134))]
        result = visual_qa.inspect(samples, step_frames=30, min_static_frames=60)
        self.assertNotIn("STATIC_PREVIEW", [issue["code"] for issue in result["issues"]])


if __name__ == "__main__":
    unittest.main()
