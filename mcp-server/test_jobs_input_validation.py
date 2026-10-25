"""Regression tests for export/job identifiers and idempotency keys."""
import unittest

import jobs


class ExportJobInputValidationTests(unittest.TestCase):
    def test_job_id_requires_full_match(self):
        valid = "job_abcdefgh"
        self.assertTrue(jobs.job_id_ok(valid))
        self.assertTrue(jobs.job_id_ok("job_" + "a" * 64))
        for invalid in (valid + "\n", valid + "\r", valid + "\x00",
                        valid + " ", "job_" + "a" * 65):
            with self.subTest(job_id=repr(invalid)):
                self.assertFalse(jobs.job_id_ok(invalid))

    def test_export_idempotency_key_rejects_control_characters(self):
        for char in ("\x00", "\t", "\n", "\r", "\x1f", "\x7f"):
            with self.subTest(char=repr(char)), self.assertRaises(ValueError):
                jobs.validate_export_request({
                    "path": "C:/Videos/final.mp4",
                    "idempotency_key": "run" + char + "1",
                })

    def test_valid_export_keys_and_length_boundaries(self):
        for key in ("run-1", "a", "x" * 128):
            with self.subTest(key_length=len(key)):
                self.assertEqual(
                    jobs.validate_export_request({
                        "path": "C:/Videos/final.mp4",
                        "idempotency_key": key,
                    })["idempotency_key"], key)
        with self.assertRaises(ValueError):
            jobs.validate_export_request({
                "path": "C:/Videos/final.mp4",
                "idempotency_key": "x" * 129,
            })


if __name__ == "__main__":
    unittest.main()
