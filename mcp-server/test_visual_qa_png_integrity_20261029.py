"""Reject PNG container corruption before accepting preview pixels as QA evidence."""
import base64
from io import BytesIO
import unittest
import zlib

from PIL import Image

import visual_qa


def _png():
    output = BytesIO()
    Image.new("RGB", (80, 45), (40, 80, 120)).save(output, format="PNG")
    return output.getvalue()


def _chunk(kind, payload):
    return (len(payload).to_bytes(4, "big") + kind + payload
            + zlib.crc32(kind + payload).to_bytes(4, "big"))


def _decode(raw):
    return visual_qa.thumbnail(base64.b64encode(raw).decode("ascii"))


class PNGIntegrityTests(unittest.TestCase):
    def test_valid_png_still_decodes(self):
        self.assertEqual(len(_decode(_png())), 64 * 36 * 3)

    def test_corrupt_ancillary_crc_is_rejected(self):
        raw = _png()
        ancillary = bytearray(_chunk(b"tEXt", b"Comment\x00safe"))
        ancillary[-1] ^= 1
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(raw[:33] + ancillary + raw[33:])

    def test_bytes_after_iend_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(_png() + b"arbitrary trailing payload")

    def test_truncated_final_chunk_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(_png()[:-1])

    def test_duplicate_ihdr_is_rejected(self):
        raw = _png()
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(raw[:33] + raw[8:33] + raw[33:])

    def test_missing_image_data_is_rejected(self):
        raw = _png()
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(raw[:33] + _chunk(b"IEND", b""))

    def test_nonconsecutive_idat_is_rejected(self):
        raw = _png()
        idat_length = int.from_bytes(raw[33:37], "big")
        self.assertEqual(raw[37:41], b"IDAT")
        compressed = raw[41:41 + idat_length]
        altered = (raw[:33] + _chunk(b"IDAT", compressed[:4])
                   + _chunk(b"tEXt", b"Comment\x00safe")
                   + _chunk(b"IDAT", compressed[4:]) + _chunk(b"IEND", b""))
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)

    def test_excessive_ancillary_chunks_are_rejected(self):
        raw = _png()
        tiny = _chunk(b"tEXt", b"")
        altered = raw[:33] + tiny * 4096 + raw[33:]
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)


    def test_unknown_critical_chunk_is_rejected(self):
        raw = _png()
        altered = raw[:33] + _chunk(b"ABCD", b"unsupported") + raw[33:]
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)

    def test_unknown_ancillary_chunk_is_allowed(self):
        raw = _png()
        altered = raw[:33] + _chunk(b"abCD", b"metadata") + raw[33:]
        self.assertEqual(len(_decode(altered)), 64 * 36 * 3)

    def test_single_plte_before_idat_is_allowed(self):
        raw = _png()
        palette = _chunk(b"PLTE", bytes((40, 80, 120)))
        self.assertEqual(len(_decode(raw[:33] + palette + raw[33:])),
                         64 * 36 * 3)

    def test_duplicate_plte_is_rejected(self):
        raw = _png()
        palette = _chunk(b"PLTE", bytes((40, 80, 120)))
        altered = raw[:33] + palette + palette + raw[33:]
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)

    def test_plte_after_idat_is_rejected(self):
        raw = _png()
        palette = _chunk(b"PLTE", bytes((40, 80, 120)))
        altered = raw[:-12] + palette + raw[-12:]
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)

    def test_invalid_reserved_chunk_type_is_rejected(self):
        raw = _png()
        altered = raw[:33] + _chunk(b"abCd".lower(), b"invalid") + raw[33:]
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            _decode(altered)


if __name__ == "__main__":
    unittest.main()
