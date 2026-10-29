"""Regression checks for bounded PNG color/profile/ancillary metadata."""
import base64
import calendar
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa
from _icc_test_profiles import rgb_icc_profile


def png():
    output = BytesIO()
    Image.new("RGB", (8, 8), "blue").save(output, format="PNG")
    return output.getvalue()


def chunks(raw):
    offset = 8
    while offset < len(raw):
        size = int.from_bytes(raw[offset:offset + 4], "big")
        yield raw[offset + 4:offset + 8], raw[offset + 8:offset + 8 + size]
        offset += size + 12


def chunk(name, data):
    return (len(data).to_bytes(4, "big") + name + data +
            zlib.crc32(name + data).to_bytes(4, "big"))


def insert(raw, name, data, *, before=b"IDAT"):
    entries = list(chunks(raw))
    index = next(i for i, (kind, _) in enumerate(entries) if kind == before)
    entries.insert(index, (name, data))
    return raw[:8] + b"".join(chunk(kind, body) for kind, body in entries)


class PNGColorMetadataTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def valid(self, raw):
        visual_qa._validate_png_container(raw)

    def test_valid_gamma_and_chromaticity(self):
        raw = insert(png(), b"gAMA", (45455).to_bytes(4, "big"))
        self.valid(insert(raw, b"cHRM", b"\x00" * 32))

    def test_reject_gamma_zero_duplicate_wrong_size_or_late(self):
        raw = insert(png(), b"gAMA", (45455).to_bytes(4, "big"))
        for candidate in (insert(png(), b"gAMA", b"\x00" * 4),
                          insert(png(), b"gAMA", b"\x01"),
                          insert(raw, b"gAMA", b"\x00\x00\x01\x00"),
                          insert(png(), b"gAMA", b"\x00\x00\x01\x00",
                                 before=b"IEND")):
            with self.subTest(candidate=candidate[:60]):
                self.invalid(candidate)

    def test_reject_chromaticity_duplicate_wrong_size_or_late(self):
        raw = insert(png(), b"cHRM", b"\x00" * 32)
        for candidate in (insert(raw, b"cHRM", b"\x00" * 32),
                          insert(png(), b"cHRM", b"\x00" * 31),
                          insert(png(), b"cHRM", b"\x00" * 32,
                                 before=b"IEND")):
            with self.subTest(candidate=candidate[:60]):
                self.invalid(candidate)

    def test_srgb_intent_and_duplicate(self):
        self.valid(insert(png(), b"sRGB", b"\x03"))
        for candidate in (insert(png(), b"sRGB", b"\x04"),
                          insert(png(), b"sRGB", b""),
                          insert(insert(png(), b"sRGB", b"\x00"),
                                 b"sRGB", b"\x01"),
                          insert(png(), b"sRGB", b"\x00", before=b"IEND")):
            with self.subTest(candidate=candidate[:60]):
                self.invalid(candidate)

    def test_valid_bounded_icc_profile(self):
        # LittleCMS-generated sRGB profile includes real required tag payloads.
        self.valid(insert(png(), b"iCCP",
                          b"Display profile\\x00\\x00" +
                          zlib.compress(rgb_icc_profile())))

    def test_reject_icc_missing_keyword_method_or_compression(self):
        for payload in (b"\x00\x00" + zlib.compress(b"ICC"),
                        b" profile\x00\x00" + zlib.compress(b"ICC"),
                        b"a  b\x00\x00" + zlib.compress(b"ICC"),
                        b"profile\x00\x01" + zlib.compress(b"ICC"),
                        b"profile\x00\x00garbage",
                        b"profile\x00\x00" + zlib.compress(b"ICC") + b"junk"):
            with self.subTest(payload=payload[:40]):
                self.invalid(insert(png(), b"iCCP", payload))

    def test_reject_icc_decompression_bomb(self):
        self.invalid(insert(png(), b"iCCP", b"profile\x00\x00" +
                            zlib.compress(b"X" * 2_000_100)))

    def test_reject_duplicate_icc_and_srgb_conflict(self):
        iccp = b"profile\x00\x00" + zlib.compress(b"ICC")
        self.invalid(insert(insert(png(), b"iCCP", iccp), b"iCCP", iccp))
        self.invalid(insert(insert(png(), b"iCCP", iccp), b"sRGB", b"\x00"))
        self.invalid(insert(insert(png(), b"sRGB", b"\x00"), b"iCCP", iccp))

    def test_reject_late_icc(self):
        self.invalid(insert(png(), b"iCCP",
                            b"profile\x00\x00" + zlib.compress(b"ICC"),
                            before=b"IEND"))

    def test_valid_physical_pixel_metadata(self):
        self.valid(insert(png(), b"pHYs",
                          (3780).to_bytes(4, "big") * 2 + b"\x01"))

    def test_reject_physical_pixel_metadata_invalid_unit_size_or_order(self):
        valid = (3780).to_bytes(4, "big") * 2 + b"\x01"
        for candidate in (insert(png(), b"pHYs", valid[:-1]),
                          insert(png(), b"pHYs", valid[:-1] + b"\x02"),
                          insert(insert(png(), b"pHYs", valid), b"pHYs", valid),
                          insert(png(), b"pHYs", valid, before=b"IEND")):
            with self.subTest(candidate=candidate[:60]):
                self.invalid(candidate)

    def test_valid_leap_day_and_leap_second(self):
        stamp = (2024).to_bytes(2, "big") + bytes([2, 29, 23, 59, 60])
        self.valid(insert(png(), b"tIME", stamp, before=b"IEND"))

    def test_reject_invalid_time_and_duplicate(self):
        for stamp in ((0, 1, 1, 0, 0, 0), (2023, 2, 29, 0, 0, 0),
                      (2024, 13, 1, 0, 0, 0), (2024, 1, 0, 0, 0, 0),
                      (2024, 1, 1, 24, 0, 0), (2024, 1, 1, 0, 60, 0),
                      (2024, 1, 1, 0, 0, 61)):
            raw = stamp[0].to_bytes(2, "big") + bytes(stamp[1:])
            with self.subTest(stamp=stamp):
                self.invalid(insert(png(), b"tIME", raw))
        raw = insert(png(), b"tIME", b"\x07\xe8\x01\x01\x00\x00\x00")
        self.invalid(insert(raw, b"tIME", b"\x07\xe8\x01\x01\x00\x00\x00"))
        self.invalid(insert(png(), b"tIME", b"\x00"))

    def test_valid_multiple_distinct_suggested_palettes(self):
        raw = insert(png(), b"sPLT", b"first\x00\x08" + b"\x00" * 6)
        self.valid(insert(raw, b"sPLT", b"second\x00\x10" + b"\x00" * 10))

    def test_reject_suggested_palette_invalid_name_depth_or_entries(self):
        for payload in (b"\x00\x08" + b"\x00" * 6,
                        b" leading\x00\x08" + b"\x00" * 6,
                        b"bad  name\x00\x08" + b"\x00" * 6,
                        b"name\x00\x07" + b"\x00" * 6,
                        b"name\x00\x08",
                        b"name\x00\x08" + b"\x00" * 5,
                        b"name\x00\x10" + b"\x00" * 9):
            with self.subTest(payload=payload):
                self.invalid(insert(png(), b"sPLT", payload))

    def test_reject_duplicate_and_late_suggested_palette(self):
        entry = b"palette\x00\x08" + b"\x00" * 6
        self.invalid(insert(insert(png(), b"sPLT", entry), b"sPLT", entry))
        self.invalid(insert(png(), b"sPLT", entry, before=b"IEND"))


if __name__ == "__main__":
    unittest.main()
