"""Visual QA must not treat invisible RGB data as visible pixels."""
import base64
from io import BytesIO
import unittest

from PIL import Image

import visual_qa


def encoded_png(image):
    output = BytesIO()
    image.save(output, format="PNG")
    return base64.b64encode(output.getvalue()).decode("ascii")


class TransparentPreviewTests(unittest.TestCase):
    def test_transparent_rgba_png_is_rejected(self):
        image = Image.new("RGBA", (80, 45), (0, 0, 0, 0))
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(encoded_png(image))

    def test_paletted_png_transparency_is_rejected(self):
        image = Image.new("P", (80, 45), 0)
        image.putpalette([0, 0, 0, 255, 255, 255])
        image.info["transparency"] = 0
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(encoded_png(image))

    def test_opaque_rgba_png_is_accepted(self):
        image = Image.new("RGBA", (80, 45), (255, 255, 255, 255))
        self.assertEqual(visual_qa.thumbnail(encoded_png(image)),
                         bytes([255] * (64 * 36 * 3)))


if __name__ == "__main__":
    unittest.main()
