"""Regression tests for complete and compact FFmpeg loudnorm diagnostics (#61)."""
import json
import unittest

import decoded_media_qa as qa


MEASUREMENT = {
    "input_i": "-20.00",
    "input_tp": "-4.00",
    "input_lra": "1.00",
    "input_thresh": "-30.00",
    "output_i": "-16.00",
}


def multiline():
    return json.dumps(MEASUREMENT, indent=2)


class LoudnormCandidateIntegrityTests(unittest.TestCase):
    def test_valid_compact_measurement(self):
        self.assertEqual((-20.0, -4.0),
                         qa._loudnorm_measurements(json.dumps(MEASUREMENT)))

    def test_compact_measurement_with_ffmpeg_prefix_log(self):
        log = "[Parsed_loudnorm_0 @ 0x123] measurements below\\n"
        log += json.dumps(MEASUREMENT) + "\\n"
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_compact_invalid_optional_measurement_rejected(self):
        record = {**MEASUREMENT, "output_tp": "Infinity"}
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(json.dumps(record))

    def test_valid_measurement_plus_truncated_second_rejected(self):
        log = multiline() + '\\n{\\n  "input_i": "-10.00",\\n'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_truncated_compact_second_rejected(self):
        log = multiline() + '\\n{"input_i": "-10.00",'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_orphaned_auxiliary_field_rejected(self):
        log = multiline() + '\\n{\\n  "output_tp": "-1.50"'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_malformed_orphaned_field_rejected(self):
        log = multiline() + '\\n"normalization_type" = "dynamic"'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_unrelated_log_lines_are_still_accepted(self):
        log = "FFmpeg initialization\\n" + multiline() + "\\nfinished"
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_multiple_complete_measurements_still_rejected(self):
        log = multiline() + "\\n" + json.dumps(MEASUREMENT)
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)


if __name__ == "__main__":
    unittest.main()
