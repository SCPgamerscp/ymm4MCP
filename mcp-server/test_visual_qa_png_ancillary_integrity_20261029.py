"""Ancillary PNG chunks must be coherent before preview QA trusts pixels."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def chunk(kind, data):
    return (len(data).to_bytes(4, "big") + kind + data +
            zlib.crc32(kind + data).to_bytes(4, "big"))


def chunks(raw):
    offset = 8
    while offset < len(raw):
        size = int.from_bytes(raw[offset:offset + 4], "big")
        yield raw[offset + 4:offset + 8], raw[offset + 8:offset + 8 + size]
        offset += size + 12


def png(mode="RGB"):
    output = BytesIO()
    image = Image.new(mode, (8, 8), 0)
    if mode == "P":
        image.putpalette([0, 0, 0, 255, 255, 255] + [0] * 762)
    image.save(output, format="PNG")
    return output.getvalue()


def insert(raw, kind, data, *, before=b"IDAT"):
    entries = list(chunks(raw))
    index = next(i for i, (name, _) in enumerate(entries) if name == before)
    entries.insert(index, (kind, data))
    return raw[:8] + b"".join(chunk(name, payload) for name, payload in entries)


class PNGAncillaryIntegrityTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def valid(self, raw):
        visual_qa._validate_png_container(raw)

    def test_valid_indexed_transparency_and_histogram(self):
        raw = png("P")
        palette_size = next(len(data) // 3 for name, data in chunks(raw)
                            if name == b"PLTE")
        raw = insert(raw, b"tRNS", b"\xff\x00")
        raw = insert(raw, b"hIST", b"\x00\x01" * palette_size)
        self.valid(raw)

    def test_valid_grayscale_transparency(self):
        self.valid(insert(png("L"), b"tRNS", b"\x00\xff"))

    def test_reject_duplicate_and_late_transparency(self):
        raw = insert(png(), b"tRNS", b"\x00\x01" * 3)
        self.invalid(insert(raw, b"tRNS", b"\x00\x02" * 3))
        self.invalid(insert(png(), b"tRNS", b"\x00\x01" * 3,
                            before=b"IEND"))

    def test_reject_invalid_transparency_color_types_and_lengths(self):
        for raw in (insert(png("RGBA"), b"tRNS", b"\x00\x00"),
                    insert(png("L"), b"tRNS", b"\x00"),
                    insert(png("P"), b"tRNS", b"\xff" * 257),
                    insert(png("L"), b"tRNS", b"\x01\x00")):
            with self.subTest(raw=raw[:40]):
                self.invalid(raw)

    def test_reject_indexed_transparency_before_palette(self):
        self.invalid(insert(png("P"), b"tRNS", b"\xff", before=b"PLTE"))

    def test_reject_background_bad_length_or_palette_index(self):
        self.invalid(insert(png(), b"bKGD", b"\x00"))
        # Pillow normally writes a 256-entry palette: index 255 is valid.
        # Reduce the PLTE to two entries to exercise an out-of-range index.
        indexed = png("P")
        small_palette = indexed[:8] + b"".join(
            chunk(kind, data[:6] if kind == b"PLTE" else data)
            for kind, data in chunks(indexed))
        self.invalid(insert(small_palette, b"bKGD", b"\x02"))
        self.invalid(insert(png("L"), b"bKGD", b"\x01\x00"))

    def test_reject_duplicate_and_late_background(self):
        raw = insert(png(), b"bKGD", b"\x00\x00" * 3)
        self.invalid(insert(raw, b"bKGD", b"\x00\x00" * 3))
        self.invalid(insert(png(), b"bKGD", b"\x00\x00" * 3,
                            before=b"IEND"))

    def test_valid_backgrounds(self):
        self.valid(insert(png(), b"bKGD", b"\x00\x00" * 3))
        self.valid(insert(png("P"), b"bKGD", b"\x00"))

    def test_reject_invalid_significant_bits(self):
        for data in (b"\x08\x08", b"\x00\x08\x08", b"\x09\x08\x08"):
            with self.subTest(data=data):
                self.invalid(insert(png(), b"sBIT", data))

    def test_reject_duplicate_and_misordered_significant_bits(self):
        raw = insert(png(), b"sBIT", b"\x08\x08\x08")
        self.invalid(insert(raw, b"sBIT", b"\x08\x08\x08"))
        self.invalid(insert(png("P"), b"sBIT", b"\x08\x08\x08"))

    def test_valid_significant_bits(self):
        self.valid(insert(png(), b"sBIT", b"\x08\x08\x08"))
        self.valid(insert(png("P"), b"sBIT", b"\x08\x08\x08",
                          before=b"PLTE"))

    def test_reject_histogram_without_palette_or_wrong_length(self):
        self.invalid(insert(png(), b"hIST", b"\x00\x01"))
        self.invalid(insert(png("P"), b"hIST", b"\x00\x01"))
        self.invalid(insert(png("P"), b"hIST", b"\x00\x01",
                            before=b"PLTE"))

    def test_reject_duplicate_and_late_histogram(self):
        raw = png("P")
        n = next(len(data) // 3 for name, data in chunks(raw)
                 if name == b"PLTE")
        raw = insert(raw, b"hIST", b"\x00\x01" * n)
        self.invalid(insert(raw, b"hIST", b"\x00\x01" * n))
        self.invalid(insert(png("P"), b"hIST", b"\x00\x01" * n,
                            before=b"IEND"))

    def test_reject_palette_after_transparency_or_background(self):
        self.invalid(insert(png("P"), b"tRNS", b"\xff", before=b"PLTE"))
        self.invalid(insert(png("P"), b"bKGD", b"\x00", before=b"PLTE"))


if __name__ == "__main__":
    unittest.main()
