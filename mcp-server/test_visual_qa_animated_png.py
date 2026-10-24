"""Preview QA must not silently inspect only the first frame of an APNG."""
import base64
from io import BytesIO
import unittest

from PIL import Image

import visual_qa


def png_base64(*, animated: bool) -> str:
    buffer = BytesIO()
    first = Image.new("RGB", (80, 45), "black")
    if animated:
        second = Image.new("RGB", (80, 45), "white")
        first.save(buffer, format="PNG", save_all=True, append_images=[second],
                   duration=[100, 100], loop=0)
    else:
        first.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("ascii")


class SingleFramePreviewTests(unittest.TestCase):
    def test_animated_png_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid preview PNG"):
            visual_qa.thumbnail(png_base64(animated=True))

    def test_normal_png_remains_accepted(self):
        thumbnail = visual_qa.thumbnail(png_base64(animated=False))
        self.assertEqual(len(thumbnail), 64 * 36 * 3)
        self.assertEqual(thumbnail, bytes(len(thumbnail)))


if __name__ == "__main__":
    unittest.main()
