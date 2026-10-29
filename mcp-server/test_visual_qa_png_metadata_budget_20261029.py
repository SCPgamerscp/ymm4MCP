"""Regression tests for cumulative PNG metadata expansion limits."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def _png():
    output = BytesIO()
    Image.new("RGB", (8, 8), "blue").save(output, format="PNG")
    return output.getvalue()


def _chunk(name, payload):
    return (len(payload).to_bytes(4, "big") + name + payload +
            zlib.crc32(name + payload).to_bytes(4, "big"))


def _insert(raw, name, payload):
    return raw[:33] + _chunk(name, payload) + raw[33:]


class CumulativeMetadataBudgetTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def test_multiple_compressed_text_chunks_exceed_total_budget(self):
        raw = _png()
        payload = b"Comment\x00\x00" + zlib.compress(b"A" * 1_100_000)
        self.invalid(_insert(_insert(raw, b"zTXt", payload), b"zTXt", payload))

    def test_mixed_plain_and_compressed_text_exceeds_budget(self):
        raw = _insert(_png(), b"tEXt", b"Author\x00" + b"A" * 1_100_000)
        self.invalid(_insert(raw, b"zTXt",
                             b"Comment\x00\x00" + zlib.compress(b"B" * 1_100_000)))

    def test_mixed_international_and_icc_profile_exceeds_budget(self):
        raw = _insert(_png(), b"iTXt",
                      b"Title\x00\x01\x00ja\x00\x00" +
                      zlib.compress(b"A" * 1_100_000))
        self.invalid(_insert(raw, b"iCCP",
                             b"Profile\x00\x00" + zlib.compress(b"B" * 1_100_000)))

    def test_multiple_uncompressed_text_chunks_exceed_budget(self):
        raw = _insert(_png(), b"tEXt", b"One\x00" + b"A" * 1_100_000)
        self.invalid(_insert(raw, b"iTXt",
                             b"Title\x00\x00\x00ja\x00\x00" + b"B" * 1_100_000))

    def test_combined_metadata_within_budget_is_accepted(self):
        raw = _insert(_png(), b"tEXt", b"Author\x00" + b"A" * 400_000)
        raw = _insert(raw, b"zTXt",
                      b"Comment\x00\x00" + zlib.compress(b"B" * 400_000))
        raw = _insert(raw, b"iTXt",
                      b"Title\x00\x01\x00ja\x00\x00" +
                      zlib.compress(b"C" * 400_000))
        raw = _insert(raw, b"iCCP",
                      b"Profile\x00\x00" + zlib.compress(b"D" * 400_000))
        visual_qa._validate_png_container(raw)


if __name__ == "__main__":
    unittest.main()
