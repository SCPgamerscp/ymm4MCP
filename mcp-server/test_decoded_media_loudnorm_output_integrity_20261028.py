"""Regression coverage for optional loudnorm fields and JSON integrity (#61)."""
import json
import unittest

import decoded_media_qa as qa


BASE = {
    "input_i": "-20.00",
    "input_tp": "-4.00",
    "input_lra": "1.00",
    "input_thresh": "-30.00",
    "output_i": "-16.00",
}


def log_with(**fields):
    return json.dumps({**BASE, **fields}, indent=2)


class LoudnormOutputIntegrityTests(unittest.TestCase):
    def test_complete_ffmpeg_first_pass_measurement(self):
        log = log_with(output_tp="-1.50", output_lra="1.20",
                       output_thresh="-26.00", target_offset="0.25",
                       normalization_type="dynamic")
        self.assertEqual((-20.0, -4.0), qa._loudnorm_measurements(log))

    def test_legacy_minimal_measurement_remains_accepted(self):
        self.assertEqual((-20.0, -4.0),
                         qa._loudnorm_measurements(log_with()))

    def test_invalid_optional_output_metrics_are_rejected(self):
        invalid = {
            "output_tp": ("NaN", "31", -1.5),
            "output_lra": ("-0.1", "101", "inf"),
            "output_thresh": ("-101", "Infinity", "bad"),
            "target_offset": ("101", "-101", "nan"),
        }
        for field, values in invalid.items():
            for value in values:
                with self.subTest(field=field, value=value):
                    with self.assertRaises(ValueError):
                        qa._loudnorm_measurements(log_with(**{field: value}))

    def test_invalid_normalization_type_is_rejected(self):
        for value in ("unknown", "LINEAR", "", None, 3, ["linear"]):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    qa._loudnorm_measurements(log_with(normalization_type=value))

    def test_overflowed_number_in_unused_metadata_is_rejected(self):
        # JSON permits numeric exponent notation, but Python can parse this
        # out-of-range value as infinity without parse_float validation.
        log = log_with(unused_metric=0).replace(
            '"unused_metric": 0', '"unused_metric": 1e999')
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)

    def test_deeply_nested_unused_metadata_is_rejected(self):
        value = 0
        for _ in range(70):
            value = {"child": value}
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log_with(unused_metric=value))


    def test_parser_recursion_limit_is_a_validation_error(self):
        # A pathological nested JSON value may overflow Python's JSON
        # decoder stack before the explicit metadata-depth check runs.
        prefix = log_with(unused_metric=0).split('"unused_metric": 0')[0] + '"unused_metric": '
        log = prefix + "[" * 1100 + "0" + "]" * 1100 + "\n}"
        with self.assertRaises(ValueError):
            qa._loudnorm_measurements(log)


if __name__ == "__main__":
    unittest.main()
