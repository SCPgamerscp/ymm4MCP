"""Malformed preview Base64 must fail through the public QA validation contract."""
import base64
from io import BytesIO
import unittest

from PIL import Image

from visual_qa import thumbnail


class PreviewBase64ValidationTests(unittest.TestCase):
    def test_invalid_alphabet_is_a_value_error(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            thumbnail("%%%not-base64%%%")

    def test_incorrect_padding_is_a_value_error(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            thumbnail("abc")

    def test_invalid_length_is_a_value_error(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            thumbnail("AAAAA")

    def test_valid_png_still_produces_rgb_thumbnail(self):
        output = BytesIO()
        Image.new("RGB", (80, 45), (10, 20, 30)).save(output, format="PNG")
        pixels = thumbnail(base64.b64encode(output.getvalue()).decode("ascii"))
        self.assertEqual(len(pixels), 64 * 36 * 3)
        self.assertEqual(pixels[:3], bytes((10, 20, 30)))

    def test_base64_with_non_png_data_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            thumbnail(base64.b64encode(b"not a png").decode("ascii"))


if __name__ == "__main__":
    unittest.main()
