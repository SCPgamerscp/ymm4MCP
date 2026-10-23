"""Regression tests for bounded PNG preview decoding."""

import base64
from io import BytesIO
import unittest
from unittest.mock import patch

from PIL import Image

import visual_qa


def png(color="white"):
    output = BytesIO()
    Image.new("RGB", (80, 45), color).save(output, format="PNG")
    return base64.b64encode(output.getvalue()).decode("ascii")


class PreviewThumbnailTests(unittest.TestCase):
    def test_valid_preview_is_resized(self):
        self.assertEqual(len(visual_qa.thumbnail(png())), 64 * 36 * 3)

    def test_pillow_decompression_bomb_is_normalized_to_value_error(self):
        # A low test-only threshold simulates a large PNG without allocating one.
        with patch.object(Image, "MAX_IMAGE_PIXELS", 100):
            with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
                visual_qa.thumbnail(png())


if __name__ == "__main__":
    unittest.main()
