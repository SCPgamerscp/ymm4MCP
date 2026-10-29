"""PNG text metadata must be syntactically valid and decompression-bounded."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def png():
    output = BytesIO()
    Image.new("RGB", (8, 8), "blue").save(output, format="PNG")
    return output.getvalue()


def chunk(name, data):
    return (len(data).to_bytes(4, "big") + name + data +
            zlib.crc32(name + data).to_bytes(4, "big"))


def insert(raw, name, payload):
    return raw[:33] + chunk(name, payload) + raw[33:]


class PNGTextMetadataTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def valid(self, raw):
        visual_qa._validate_png_container(raw)

    def test_valid_latin1_text(self):
        self.valid(insert(png(), b"tEXt", b"Author\x00Name \xe9"))

    def test_reject_latin1_text_without_separator_or_keyword(self):
        for payload in (b"no separator", b"\x00text",
                        b" leading\x00text", b"bad  key\x00text",
                        b"a" * 80 + b"\x00text"):
            with self.subTest(payload=payload[:30]):
                self.invalid(insert(png(), b"tEXt", payload))

    def test_reject_oversized_latin1_text(self):
        self.invalid(insert(png(), b"tEXt", b"key\x00" + b"x" * 2_000_001))

    def test_valid_compressed_latin1_text(self):
        self.valid(insert(png(), b"zTXt", b"Comment\x00\x00" +
                          zlib.compress(b"hello" * 100)))

    def test_reject_compressed_text_invalid_method_or_stream(self):
        for payload in (b"key\x00\x01" + zlib.compress(b"text"),
                        b"key\x00\x00broken",
                        b"key\x00\x00" + zlib.compress(b"text") + b"junk",
                        b"key\x00\x00"):
            with self.subTest(payload=payload[:40]):
                self.invalid(insert(png(), b"zTXt", payload))

    def test_reject_compressed_text_decompression_bomb(self):
        self.invalid(insert(png(), b"zTXt", b"key\x00\x00" +
                            zlib.compress(b"x" * 2_000_100)))

    def test_valid_uncompressed_international_text(self):
        self.valid(insert(png(), b"iTXt",
                          b"Title\x00\x00\x00ja-JP\x00" +
                          "題名".encode("utf-8") + b"\x00" +
                          "日本語の文章".encode("utf-8")))

    def test_valid_compressed_international_text(self):
        self.valid(insert(png(), b"iTXt",
                          b"Title\x00\x01\x00ja\x00\x00" +
                          zlib.compress("こんにちは".encode("utf-8"))))

    def test_reject_international_text_header_and_missing_delimiters(self):
        for payload in (b"Title\x00\x02\x00ja\x00\x00hello",
                        b"Title\x00\x01\x01ja\x00\x00hello",
                        b"Title\x00\x00\x00ja",
                        b"Title\x00\x00\x00ja\x00missing-translated-delimiter",
                        b"\x00\x00\x00ja\x00\x00hello",
                        b"Title\x00\x00\x00ja_US\x00\x00hello"):
            with self.subTest(payload=payload):
                self.invalid(insert(png(), b"iTXt", payload))

    def test_reject_international_text_invalid_utf8(self):
        for payload in (b"Title\x00\x00\x00ja\x00\xff\x00hello",
                        b"Title\x00\x00\x00ja\x00\x00\xff",
                        b"Title\x00\x01\x00ja\x00\x00" +
                        zlib.compress(b"\xff")):
            with self.subTest(payload=payload[:40]):
                self.invalid(insert(png(), b"iTXt", payload))

    def test_reject_international_text_bad_compression_and_bomb(self):
        for payload in (b"Title\x00\x01\x00ja\x00\x00invalid",
                        b"Title\x00\x01\x00ja\x00\x00" +
                        zlib.compress(b"X" * 2_000_100),
                        b"Title\x00\x01\x00ja\x00\x00" +
                        zlib.compress(b"ok") + b"trailing"):
            with self.subTest(payload=payload[:40]):
                self.invalid(insert(png(), b"iTXt", payload))


if __name__ == "__main__":
    unittest.main()
