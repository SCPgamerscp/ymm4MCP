"""PNG critical-header and palette semantics must be checked before preview QA."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def chunk(kind, data):
    return (len(data).to_bytes(4, "big") + kind + data +
            zlib.crc32(kind + data).to_bytes(4, "big"))


def png(mode="RGB"):
    output = BytesIO()
    image = Image.new(mode, (8, 8), 0)
    if mode == "P":
        image.putpalette([0, 0, 0, 255, 255, 255] + [0] * (768 - 6))
    image.save(output, format="PNG")
    return output.getvalue()


def with_ihdr(raw, **fields):
    header = bytearray(raw[16:29])
    positions = {"bit_depth": 8, "color_type": 9, "compression": 10,
                 "filtering": 11, "interlace": 12}
    for key, value in fields.items():
        if key in positions:
            header[positions[key]] = value
        else:
            header[0 if key == "width" else 4:
                   4 if key == "width" else 8] = value.to_bytes(4, "big")
    return raw[:8] + chunk(b"IHDR", bytes(header)) + raw[33:]


def chunks(raw):
    offset = 8
    while offset < len(raw):
        length = int.from_bytes(raw[offset:offset + 4], "big")
        kind = raw[offset + 4:offset + 8]
        yield kind, raw[offset + 8:offset + 8 + length]
        offset += 12 + length


class PNGSemanticsTests(unittest.TestCase):
    def assert_invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def test_valid_rgb_grayscale_and_indexed_previews(self):
        for mode in ("RGB", "L", "P"):
            with self.subTest(mode=mode):
                raw = png(mode)
                self.assertEqual(len(visual_qa.thumbnail(
                    base64.b64encode(raw).decode("ascii"))), 64 * 36 * 3)

    def test_zero_dimensions_rejected(self):
        for key in ("width", "height"):
            with self.subTest(key=key):
                self.assert_invalid(with_ihdr(png(), **{key: 0}))

    def test_illegal_color_type_and_bit_depth_pairs_rejected(self):
        for color, depth in ((1, 8), (5, 8), (2, 4), (3, 16), (6, 1), (4, 2)):
            with self.subTest(color=color, depth=depth):
                self.assert_invalid(with_ihdr(png(), color_type=color, bit_depth=depth))

    def test_invalid_compression_filter_interlace_rejected(self):
        for key, value in (("compression", 1), ("filtering", 1), ("interlace", 2)):
            with self.subTest(key=key):
                self.assert_invalid(with_ihdr(png(), **{key: value}))

    def test_indexed_png_requires_palette(self):
        raw = png("P")
        self.assert_invalid(raw[:8] + b"".join(
            chunk(kind, data) for kind, data in chunks(raw) if kind != b"PLTE"))

    def test_grayscale_png_forbids_palette(self):
        raw = png("L")
        self.assert_invalid(raw[:33] + chunk(b"PLTE", b"\x00\x00\x00") + raw[33:])

    def test_palette_must_be_nonempty_triplets_and_bounded(self):
        raw = png()
        for palette in (b"", b"\x00", b"\x00" * 769, b"\x00" * 770):
            with self.subTest(size=len(palette)):
                self.assert_invalid(raw[:33] + chunk(b"PLTE", palette) + raw[33:])

    def test_indexed_palette_must_fit_bit_depth(self):
        raw = png("P")
        modified = with_ihdr(raw, bit_depth=1)
        # Existing 256-entry palette exceeds the two entries allowed at depth 1.
        self.assert_invalid(modified)

    def test_empty_idat_payload_rejected(self):
        raw = png()
        altered = raw[:8] + b"".join(
            chunk(kind, b"" if kind == b"IDAT" else data)
            for kind, data in chunks(raw))
        self.assert_invalid(altered)

    def test_optional_rgb_palette_remains_valid(self):
        raw = png()
        altered = raw[:33] + chunk(b"PLTE", b"\x00\x00\x00") + raw[33:]
        self.assertEqual(len(visual_qa.thumbnail(
            base64.b64encode(altered).decode("ascii"))), 64 * 36 * 3)


if __name__ == "__main__":
    unittest.main()
