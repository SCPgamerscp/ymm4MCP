"""Regressions for strict PNG ancillary evidence validation (Issue #61)."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def png(mode="RGB"):
    out = BytesIO()
    Image.new(mode, (8, 8), "blue" if mode == "RGB" else 0).save(out, "PNG")
    return out.getvalue()


def chunk(kind, payload):
    return (len(payload).to_bytes(4, "big") + kind + payload +
            zlib.crc32(kind + payload).to_bytes(4, "big"))


def insert(raw, kind, payload, *, after_idat=False):
    if after_idat:
        end = raw.rfind(b"IEND") - 4
        return raw[:end] + chunk(kind, payload) + raw[end:]
    return raw[:33] + chunk(kind, payload) + raw[33:]


def profile(space=b"RGB ", tags=()):
    result = bytearray(132 + 12 * len(tags))
    result[:4] = len(result).to_bytes(4, "big")
    result[16:20] = space
    result[20:24] = b"XYZ "
    result[36:40] = b"acsp"
    result[128:132] = len(tags).to_bytes(4, "big")
    for i, (offset, size) in enumerate(tags):
        start = 132 + 12 * i
        result[start:start + 4] = b"desc"
        result[start + 4:start + 8] = offset.to_bytes(4, "big")
        result[start + 8:start + 12] = size.to_bytes(4, "big")
    return bytes(result)


def iccp(raw, data):
    return insert(raw, b"iCCP", b"Profile\x00\x00" +
                  zlib.compress(data))


class PNGMetadataIntegrityTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def test_valid_rgb_icc(self):
        visual_qa._validate_png_container(iccp(png(), profile()))

    def test_valid_gray_icc(self):
        visual_qa._validate_png_container(iccp(png("L"), profile(b"GRAY")))

    def test_reject_short_icc(self):
        self.invalid(iccp(png(), b"ICC" * 30))

    def test_reject_icc_declared_size_mismatch(self):
        raw = bytearray(profile())
        raw[:4] = (133).to_bytes(4, "big")
        self.invalid(iccp(png(), raw))

    def test_reject_icc_signature_mismatch(self):
        raw = bytearray(profile())
        raw[36:40] = b"fail"
        self.invalid(iccp(png(), raw))

    def test_reject_icc_color_space_mismatch(self):
        self.invalid(iccp(png(), profile(b"GRAY")))
        self.invalid(iccp(png("L"), profile(b"RGB ")))

    def test_reject_icc_connection_space_mismatch(self):
        raw = bytearray(profile())
        raw[20:24] = b"bad!"
        self.invalid(iccp(png(), raw))

    def test_reject_icc_tag_table_overflow(self):
        raw = bytearray(profile())
        raw[128:132] = (1).to_bytes(4, "big")
        self.invalid(iccp(png(), raw))

    def test_reject_icc_out_of_bounds_tag(self):
        self.invalid(iccp(png(), profile(tags=[(1024, 32)])))
        self.invalid(iccp(png(), profile(tags=[(144, 8)])))

    def test_reject_apng_control_even_for_one_frame(self):
        for kind, payload in ((b"acTL", b"\x00\x00\x00\x01\x00\x00\x00\x00"),
                              (b"fcTL", b"\x00" * 26),
                              (b"fdAT", b"\x00" * 4)):
            with self.subTest(kind=kind):
                self.invalid(insert(png(), kind, payload))

    def test_reject_apng_control_after_idat(self):
        self.invalid(insert(png(), b"fcTL", b"\x00" * 26, after_idat=True))

    def test_reject_embedded_nul_in_plain_text(self):
        self.invalid(insert(png(), b"tEXt", b"Author\x00foo\x00bar"))

    def test_reject_embedded_nul_in_compressed_text(self):
        self.invalid(insert(png(), b"zTXt", b"Author\x00\x00" +
                            zlib.compress(b"foo\x00bar")))

    def test_reject_embedded_nul_in_itxt_translation(self):
        self.invalid(insert(png(), b"iTXt",
                            b"Title\x00\x00\x00ja\x00foo\x00bar\x00text"))

    def test_reject_embedded_nul_in_itxt_plain_and_compressed(self):
        for flag, data in ((b"\x00", b"foo\x00bar"),
                           (b"\x01", zlib.compress(b"foo\x00bar"))):
            with self.subTest(flag=flag):
                self.invalid(insert(png(), b"iTXt",
                                    b"Title\x00" + flag + b"\x00ja\x00\x00" + data))

    def test_reject_invalid_itxt_language_subtags(self):
        for lang in (b"-ja", b"ja-", b"ja--JP", b"ja_JP", b"123", b"japaneseverylong"):
            with self.subTest(lang=lang):
                self.invalid(insert(png(), b"iTXt",
                                    b"Title\x00\x00\x00" + lang + b"\x00\x00text"))

    def test_accept_valid_itxt_language_subtags(self):
        for lang in (b"", b"ja-JP", b"zh-Hant-TW", b"x-private"):
            with self.subTest(lang=lang):
                visual_qa._validate_png_container(
                    insert(png(), b"iTXt",
                           b"Title\x00\x00\x00" + lang + b"\x00\x00text"))


if __name__ == "__main__":
    unittest.main()
