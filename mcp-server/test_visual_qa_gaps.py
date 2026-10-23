"""Regression coverage for discontinuous visual QA preview samples."""
import unittest
import visual_qa


class DiscontinuousSamplesTests(unittest.TestCase):
    def test_black_segments_do_not_bridge_unsampled_gap(self):
        black = bytes([0] * (64 * 36 * 3))
        result = visual_qa.inspect([(0, black), (30, black), (90, black), (120, black)],
                                   step_frames=30, min_static_frames=60)
        self.assertEqual([(i['start_frame'], i['end_frame']) for i in result['issues']
                          if i['code'] == 'BLACK_FRAME'], [(0, 30), (90, 120)])
        self.assertFalse(any(i['code'] == 'STATIC_PREVIEW' for i in result['issues']))

    def test_static_run_ends_at_last_observed_frame_before_gap(self):
        black = bytes([0] * (64 * 36 * 3))
        result = visual_qa.inspect([(0, black), (30, black), (60, black),
                                    (120, black), (150, black)], step_frames=30)
        self.assertEqual([(i['start_frame'], i['end_frame']) for i in result['issues']
                          if i['code'] == 'STATIC_PREVIEW'], [(0, 60)])

    def test_white_samples_far_apart_do_not_imply_static_video(self):
        white = bytes([255] * (64 * 36 * 3))
        result = visual_qa.inspect([(0, white), (1000, white)], step_frames=30)
        self.assertEqual([(i['start_frame'], i['end_frame']) for i in result['issues']
                          if i['code'] == 'WHITE_FRAME'], [(0, 0), (1000, 1000)])
        self.assertFalse(any(i['code'] == 'STATIC_PREVIEW' for i in result['issues']))

    def test_regular_sampling_still_detects_static(self):
        gray = bytes([128] * (64 * 36 * 3))
        result = visual_qa.inspect([(0, gray), (30, gray), (60, gray), (90, gray)], step_frames=30)
        self.assertEqual([(i['start_frame'], i['end_frame']) for i in result['issues']
                          if i['code'] == 'STATIC_PREVIEW'], [(0, 90)])


if __name__ == '__main__':
    unittest.main()
