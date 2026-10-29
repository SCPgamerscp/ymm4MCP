"""Regressions for strict PNG ancillary evidence validation (Issue #61)."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa
from _icc_test_profiles import gray_icc_profile, rgb_icc_profile


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
    result = bytearray(gray_icc_profile() if space == b"GRAY"
                       else rgb_icc_profile())
    if tags:
        # Deliberately corrupt the first genuine tag record to exercise bounds.
        offset, size = tags[0]
        result[136:144] = offset.to_bytes(4, "big") + size.to_bytes(4, "big")
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
        raw[128:132] = (9999).to_bytes(4, "big")
        self.invalid(iccp(png(), raw))

    def test_reject_icc_out_of_bounds_tag(self):
        self.invalid(iccp(png(), profile(tags=[(1024, 32)])))
        self.invalid(iccp(png(), profile(tags=[(144, 8)])))

    def test_reject_missing_required_icc_tags(self):
        original = rgb_icc_profile()
        count = int.from_bytes(original[128:132], "big")
        for missing in (b"desc", b"cprt", b"wtpt", b"rXYZ",
                        b"gXYZ", b"bXYZ", b"rTRC", b"gTRC", b"bTRC"):
            with self.subTest(missing=missing):
                raw = bytearray(original)
                for i in range(count):
                    start = 132 + 12 * i
                    if raw[start:start + 4] == missing:
                        raw[start:start + 4] = b"zzzz"
                        break
                self.invalid(iccp(png(), raw))

    def test_reject_missing_gray_tone_curve(self):
        raw = bytearray(gray_icc_profile())
        count = int.from_bytes(raw[128:132], "big")
        for i in range(count):
            start = 132 + 12 * i
            if raw[start:start + 4] == b"kTRC":
                raw[start:start + 4] = b"zzzz"
                break
        self.invalid(iccp(png("L"), raw))

    def test_reject_duplicate_icc_tag_or_invalid_required_type(self):
        raw = bytearray(rgb_icc_profile())
        raw[144:148] = b"desc"  # Duplicate first tag signature.
        self.invalid(iccp(png(), raw))
        raw = bytearray(rgb_icc_profile())
        count = int.from_bytes(raw[128:132], "big")
        for i in range(count):
            start = 132 + 12 * i
            if raw[start:start + 4] == b"wtpt":
                offset = int.from_bytes(raw[start + 4:start + 8], "big")
                raw[offset:offset + 4] = b"JUNK"
                break
        self.invalid(iccp(png(), raw))

    def test_reject_missing_palette_color_in_decoded_pixels(self):
        # Pillow keeps the index (3) but silently converts it to black even
        # though the shortened PLTE contains only indices 0 and 1.
        image = Image.new("P", (8, 8), 3)
        image.putpalette([0, 0, 0, 255, 255, 255] + [0] * 762)
        out = BytesIO()
        image.save(out, "PNG")
        raw = out.getvalue()
        offset = 8
        rebuilt = bytearray(raw[:8])
        while offset < len(raw):
            size = int.from_bytes(raw[offset:offset + 4], "big")
            kind = raw[offset + 4:offset + 8]
            payload = raw[offset + 8:offset + 8 + size]
            if kind == b"PLTE":
                payload = payload[:6]
            rebuilt.extend(chunk(kind, payload))
            offset += 12 + size
        raw = bytes(rebuilt)
        visual_qa._validate_png_container(raw)  # Structurally valid PNG.
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def test_accept_indexed_pixels_within_short_palette(self):
        image = Image.new("P", (8, 8), 1)
        image.putpalette([0, 0, 0, 255, 255, 255])
        out = BytesIO()
        image.save(out, "PNG")
        raw = out.getvalue()
        self.assertEqual(len(visual_qa.thumbnail(
            base64.b64encode(raw).decode("ascii"))), 64 * 36 * 3)

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
