"""Regression coverage for ISO BMFF ftyp box-size validation."""
import tempfile
import unittest
from pathlib import Path

import jobs


class Mp4BoxSizeTests(unittest.TestCase):
    def _verify(self, data: bytes) -> dict:
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "export.mp4"
            target.write_bytes(data.ljust(40, b"\0"))
            return jobs.verify_media_file(str(target), "mp4")

    def test_accepts_regular_and_extended_size_ftyp(self):
        standard = (16).to_bytes(4, "big") + b"ftypisom" + b"\0" * 4
        result = self._verify(standard)
        self.assertTrue(result["ok"])
        self.assertEqual(result["brand"], "isom")

        extended = (1).to_bytes(4, "big") + b"ftyp" + (24).to_bytes(8, "big") + b"mp42" + b"\0" * 4
        result = self._verify(extended)
        self.assertTrue(result["ok"])
        self.assertEqual(result["brand"], "mp42")

    def test_rejects_short_zero_and_oversized_ftyp_boxes(self):
        for declared_size in (0, 8, 12, 15, 41, 0xFFFFFFFF):
            with self.subTest(declared_size=declared_size):
                data = declared_size.to_bytes(4, "big") + b"ftypisom" + b"\0" * 4
                result = self._verify(data)
                self.assertFalse(result["ok"])
                self.assertEqual(result["error_code"], "EXPORT_VERIFY_FAILED")

    def test_rejects_invalid_extended_size(self):
        for declared_size in (0, 8, 16, 23, 41, 0xFFFFFFFFFFFFFFFF):
            with self.subTest(declared_size=declared_size):
                data = (1).to_bytes(4, "big") + b"ftyp" + declared_size.to_bytes(8, "big") + b"isom" + b"\0" * 4
                result = self._verify(data)
                self.assertFalse(result["ok"])
                self.assertEqual(result["error_code"], "EXPORT_VERIFY_FAILED")


if __name__ == "__main__":
    unittest.main()
