"""Regression tests for Windows FFmpeg loudnorm logs and disguised invalid blocks."""
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


class LoudnormWindowsAndDisguisedCandidatesTests(unittest.TestCase):
    def test_multiline_crlf_measurement_is_accepted(self):
        log = json.dumps(MEASUREMENT, indent=2).replace("\n", "\r\n") + "\r\n"
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_compact_crlf_measurement_is_accepted(self):
        log = "FFmpeg starting\r\n" + json.dumps(MEASUREMENT) + "\r\n"
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_legacy_cr_only_measurement_is_accepted(self):
        log = json.dumps(MEASUREMENT, indent=2).replace("\n", "\r")
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_valid_crlf_followed_by_truncated_measurement_rejected(self):
        log = json.dumps(MEASUREMENT, indent=2).replace("\n", "\r\n")
        log += '\r\n{\r\n  "input_i": "-10.00",\r\n'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_unquoted_field_rejected(self):
        log = json.dumps(MEASUREMENT, indent=2) + '\n{\n  input_i: "-10.00"\n}'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_escaped_corrupt_field_rejected(self):
        log = json.dumps(MEASUREMENT, indent=2)
        log += '\n{"\\u0069nput_i": "-10.00",}'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_truncated_escaped_field_rejected(self):
        log = json.dumps(MEASUREMENT, indent=2)
        log += '\n{"\\u0069nput_i": "-10.00",'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_valid_measurement_plus_broken_separator_rejected(self):
        log = json.dumps(MEASUREMENT, indent=2)
        log += '\n{"input_i", "-10.00"}'
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_unrelated_json_still_ignored(self):
        log = json.dumps({"event": "unrelated"}) + "\n"
        log += json.dumps(MEASUREMENT)
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))


if __name__ == "__main__":
    unittest.main()
