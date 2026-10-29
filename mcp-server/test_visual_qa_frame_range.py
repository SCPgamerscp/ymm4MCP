"""Visual QA must not certify frame positions beyond YMM4's supported range."""
import unittest

from editing import MAX_FRAME
import visual_qa


GRAY = bytes([128] * (64 * 36 * 3))


class VisualQaFrameRangeTests(unittest.TestCase):
    def test_rejects_out_of_range_capture_frames(self):
        for frame in (-1, MAX_FRAME + 1, 10 ** 100):
            with self.subTest(frame=frame), self.assertRaises(ValueError):
                visual_qa.inspect([(frame, GRAY)], step_frames=30)

    def test_rejects_out_of_range_sampling_parameters(self):
        for value in (MAX_FRAME + 1, 10 ** 100):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    visual_qa.inspect([(0, GRAY)], step_frames=value)
                with self.assertRaises(ValueError):
                    visual_qa.inspect([(0, GRAY)], step_frames=30,
                                      min_static_frames=value)

    def test_boundary_frames_remain_valid(self):
        result = visual_qa.inspect([(0, GRAY), (MAX_FRAME, GRAY)],
                                   step_frames=MAX_FRAME,
                                   min_static_frames=MAX_FRAME)
        self.assertTrue(result["success"])
        self.assertEqual(result["sample_count"], 2)


if __name__ == "__main__":
    unittest.main()
