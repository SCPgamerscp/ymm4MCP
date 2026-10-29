"""Regression tests for PNG IDAT stream and filtered scanline integrity."""

import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def chunk(name, data):
    return (len(data).to_bytes(4, "big") + name + data +
            zlib.crc32(name + data).to_bytes(4, "big"))


def entries(raw):
    offset = 8
    while offset < len(raw):
        size = int.from_bytes(raw[offset:offset + 4], "big")
        yield raw[offset + 4:offset + 8], raw[offset + 8:offset + 8 + size]
        offset += size + 12


def png():
    output = BytesIO()
    Image.new("RGB", (8, 8), "blue").save(output, format="PNG")
    return output.getvalue()


def rewrite_idat(raw, compressed_parts):
    result = []
    inserted = False
    for kind, data in entries(raw):
        if kind == b"IDAT":
            if not inserted:
                result.extend((b"IDAT", part) for part in compressed_parts)
                inserted = True
        else:
            result.append((kind, data))
    return raw[:8] + b"".join(chunk(kind, data) for kind, data in result)


def decoded_pixels(raw):
    compressed = b"".join(data for kind, data in entries(raw) if kind == b"IDAT")
    return zlib.decompress(compressed)


def with_scanlines(raw, scanlines):
    return rewrite_idat(raw, [zlib.compress(scanlines)])


def adam7_gray(width=8, height=8):
    header = (width.to_bytes(4, "big") + height.to_bytes(4, "big") +
              bytes((8, 0, 0, 0, 1)))
    passes = ((0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8),
              (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2),
              (0, 1, 1, 2))
    rows = bytearray()
    for x0, y0, dx, dy in passes:
        if width > x0 and height > y0:
            columns = (width - x0 + dx - 1) // dx
            count = (height - y0 + dy - 1) // dy
            rows.extend((b"\x00" + b"\x80" * columns) * count)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) +
            chunk(b"IDAT", zlib.compress(bytes(rows))) + chunk(b"IEND", b""))


class PNGScanlineIntegrityTests(unittest.TestCase):
    def invalid(self, raw):
        with self.assertRaises(ValueError):
            visual_qa._validate_png_container(raw)
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))

    def test_regular_png_remains_valid(self):
        raw = png()
        visual_qa._validate_png_container(raw)
        self.assertEqual(len(visual_qa.thumbnail(
            base64.b64encode(raw).decode("ascii"))), 64 * 36 * 3)

    def test_idat_split_across_chunks_remains_valid(self):
        raw = png()
        compressed = b"".join(data for name, data in entries(raw)
                              if name == b"IDAT")
        middle = len(compressed) // 2
        visual_qa._validate_png_container(
            rewrite_idat(raw, [compressed[:middle], compressed[middle:]]))

    def test_invalid_scanline_filter_is_rejected(self):
        raw = png()
        pixels = bytearray(decoded_pixels(raw))
        pixels[0] = 5
        self.invalid(with_scanlines(raw, pixels))

    def test_invalid_filter_on_later_row_is_rejected(self):
        raw = png()
        pixels = bytearray(decoded_pixels(raw))
        pixels[1 + 8 * 3] = 255
        self.invalid(with_scanlines(raw, pixels))

    def test_missing_decompressed_scanline_is_rejected(self):
        raw = png()
        self.invalid(with_scanlines(raw, decoded_pixels(raw)[:-1]))

    def test_excess_decompressed_scanline_is_rejected(self):
        raw = png()
        self.invalid(with_scanlines(raw, decoded_pixels(raw) + b"\x00"))

    def test_truncated_zlib_stream_is_rejected(self):
        raw = png()
        compressed = b"".join(data for name, data in entries(raw)
                              if name == b"IDAT")
        self.invalid(rewrite_idat(raw, [compressed[:-2]]))

    def test_trailing_bytes_after_zlib_end_are_rejected(self):
        raw = png()
        compressed = b"".join(data for name, data in entries(raw)
                              if name == b"IDAT")
        self.invalid(rewrite_idat(raw, [compressed + b"junk"]))

    def test_second_zlib_stream_is_rejected(self):
        raw = png()
        compressed = b"".join(data for name, data in entries(raw)
                              if name == b"IDAT")
        self.invalid(rewrite_idat(raw, [compressed, zlib.compress(b"ignored")]))

    def test_valid_adam7_pass_layout_is_accepted(self):
        visual_qa._validate_png_container(adam7_gray())

    def test_truncated_adam7_pass_is_rejected(self):
        raw = adam7_gray()
        self.invalid(with_scanlines(raw, decoded_pixels(raw)[:-1]))

    def test_invalid_adam7_filter_is_rejected(self):
        raw = adam7_gray()
        pixels = bytearray(decoded_pixels(raw))
        pixels[0] = 7
        self.invalid(with_scanlines(raw, pixels))

    def test_excessive_pixel_dimensions_fail_before_decompression(self):
        raw = png()
        header = bytearray(raw[16:29])
        header[:4] = (32_000_001).to_bytes(4, "big")
        altered = raw[:8] + chunk(b"IHDR", header) + raw[33:]
        self.invalid(altered)


if __name__ == "__main__":
    unittest.main()
