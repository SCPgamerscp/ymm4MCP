"""sRGB previews must not carry contradictory gamma or chromaticity tags."""
from io import BytesIO
import unittest
import zlib

from PIL import Image
import visual_qa


def png():
    out = BytesIO()
    Image.new("RGB", (8, 8), "blue").save(out, format="PNG")
    return out.getvalue()


def chunk(kind, payload):
    return (len(payload).to_bytes(4, "big") + kind + payload +
            zlib.crc32(kind + payload).to_bytes(4, "big"))


def insert(raw, kind, payload):
    return raw[:33] + chunk(kind, payload) + raw[33:]


GAMMA = (45455).to_bytes(4, "big")
CHRM = b"".join(value.to_bytes(4, "big") for value in (
    31270, 32900, 64000, 33000, 30000, 60000, 15000, 6000))


class SRGBConsistencyTests(unittest.TestCase):
    def test_valid_srgb_and_canonical_metadata_in_any_order(self):
        for order in ((b"sRGB", b"gAMA", b"cHRM"),
                      (b"cHRM", b"gAMA", b"sRGB")):
            raw = png()
            for kind in order:
                raw = insert(raw, kind, {
                    b"sRGB": b"\x00", b"gAMA": GAMMA, b"cHRM": CHRM}[kind])
            with self.subTest(order=order):
                visual_qa._validate_png_container(raw)

    def test_reject_contradictory_gamma_in_either_order(self):
        bad_gamma = (50000).to_bytes(4, "big")
        for order in ((b"sRGB", b"gAMA"), (b"gAMA", b"sRGB")):
            raw = png()
            for kind in order:
                raw = insert(raw, kind,
                             b"\x00" if kind == b"sRGB" else bad_gamma)
            with self.subTest(order=order):
                with self.assertRaises(ValueError):
                    visual_qa._validate_png_container(raw)

    def test_reject_contradictory_chromaticity_in_either_order(self):
        bad_chrm = b"\x00" * 32
        for order in ((b"sRGB", b"cHRM"), (b"cHRM", b"sRGB")):
            raw = png()
            for kind in order:
                raw = insert(raw, kind,
                             b"\x00" if kind == b"sRGB" else bad_chrm)
            with self.subTest(order=order):
                with self.assertRaises(ValueError):
                    visual_qa._validate_png_container(raw)

    def test_non_srgb_metadata_remains_compatible(self):
        raw = insert(png(), b"gAMA", (50000).to_bytes(4, "big"))
        visual_qa._validate_png_container(insert(raw, b"cHRM", b"\x00" * 32))


if __name__ == "__main__":
    unittest.main()
