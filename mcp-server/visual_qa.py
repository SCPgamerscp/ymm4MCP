"""Conservative preview QA on a bounded set of PNG captures."""

import base64
import binascii
import zlib
from io import BytesIO

from PIL import Image, UnidentifiedImageError

from editing import MAX_FRAME


def _validate_png_container(raw: bytes) -> None:
    """Reject malformed PNG chunks before Pillow discards their metadata.

    Pillow can decode the pixels of a PNG even when an ancillary chunk has a
    corrupt CRC or arbitrary bytes follow IEND. Neither is a trustworthy
    single-frame preview from YMM4.
    """
    if not raw.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("invalid PNG signature")
    offset = 8
    chunk_count = 0
    seen_ihdr = seen_idat = seen_iend = idat_finished = False
    while offset < len(raw):
        chunk_count += 1
        if chunk_count > 4096 or len(raw) - offset < 12:
            raise ValueError("invalid PNG chunk count or header")
        length = int.from_bytes(raw[offset:offset + 4], "big")
        kind = raw[offset + 4:offset + 8]
        end = offset + 12 + length
        if end > len(raw):
            raise ValueError("truncated PNG chunk")
        expected_crc = int.from_bytes(raw[offset + 8 + length:end], "big")
        if zlib.crc32(raw[offset + 4:offset + 8 + length]) != expected_crc:
            raise ValueError("PNG chunk CRC mismatch")
        if not seen_ihdr:
            if kind != b"IHDR" or length != 13:
                raise ValueError("PNG must begin with IHDR")
            seen_ihdr = True
        elif kind == b"IHDR":
            raise ValueError("duplicate PNG IHDR")
        if kind == b"IDAT":
            if idat_finished:
                raise ValueError("nonconsecutive PNG IDAT chunks")
            seen_idat = True
        elif seen_idat:
            idat_finished = True
        if kind == b"IEND":
            if length != 0 or not seen_idat or end != len(raw):
                raise ValueError("invalid PNG end")
            seen_iend = True
            break
        offset = end
    if not seen_iend:
        raise ValueError("PNG is missing IEND")


def thumbnail(image_b64: str) -> bytes:
    if not isinstance(image_b64, str) or len(image_b64) > 14_000_000:
        raise ValueError("preview image is missing or too large")
    try:
        raw = base64.b64decode(image_b64, validate=True)
        if len(raw) > 10_000_000:
            raise ValueError("preview image is too large")
        _validate_png_container(raw)
        with Image.open(BytesIO(raw)) as image:
            if image.format != "PNG" or image.width * image.height > 32_000_000:
                raise ValueError("preview must be a PNG of at most 32 megapixels")
            # A capture represents one timeline frame. APNG would silently use only
            # its first animation frame and could make visual QA report false results.
            if getattr(image, "n_frames", 1) != 1:
                raise ValueError("preview must contain exactly one PNG frame")
            image.load()
            # Alpha is not represented in RGB thumbnails. Discarding it can
            # turn an invisible black/white pixel into a false QA observation.
            # Reject partially transparent captures instead of guessing which
            # YMM4 canvas color should be composited behind them.
            if "A" in image.getbands() or "transparency" in image.info:
                if image.convert("RGBA").getchannel("A").getextrema()[0] < 255:
                    raise ValueError("preview contains transparent pixels")
            return image.convert("RGB").resize((64, 36)).tobytes()
    except (ValueError, binascii.Error, UnidentifiedImageError,
            Image.DecompressionBombError, OSError) as exc:
        raise ValueError("invalid preview PNG") from exc



def _visually_unchanged(before: bytes, after: bytes) -> bool:
    """Detect small localized motion as well as broad low-contrast changes.

    Mean RGB difference alone misses a moving object that occupies only a few
    thumbnail pixels. Require at least four pixels with a substantial channel
    change to avoid treating isolated capture noise as meaningful motion.
    """
    if sum(abs(a - b) for a, b in zip(before, after)) >= len(before):
        return False
    strong_pixels = 0
    for i in range(0, len(before), 3):
        if max(abs(before[i + channel] - after[i + channel])
               for channel in range(3)) >= 32:
            strong_pixels += 1
            if strong_pixels >= 4:
                return False
    return True

def inspect(samples: list[tuple[int, bytes]], *, step_frames: int,
            min_static_frames: int = 60, black_as_error: bool = False,
            white_as_error: bool = False) -> dict:
    """Only mark observed frames; gaps between samples are never called exact boundaries."""
    if not isinstance(samples, list) or not 1 <= len(samples) <= 40:
        raise ValueError("samples must contain 1..40 captures")
    if isinstance(step_frames, bool) or not isinstance(step_frames, int) or not 1 <= step_frames <= MAX_FRAME:
        raise ValueError("step_frames must be positive")
    if isinstance(min_static_frames, bool) or not isinstance(min_static_frames, int) or not 1 <= min_static_frames <= MAX_FRAME:
        raise ValueError("min_static_frames must be positive")
    if not isinstance(black_as_error, bool) or not isinstance(white_as_error, bool):
        raise ValueError("black_as_error and white_as_error must be boolean")
    last_frame = -1
    for sample in samples:
        if (not isinstance(sample, tuple) or len(sample) != 2 or
                isinstance(sample[0], bool) or not isinstance(sample[0], int) or
                sample[0] <= last_frame or sample[0] > MAX_FRAME or not isinstance(sample[1], bytes) or
                len(sample[1]) != 64 * 36 * 3):
            raise ValueError("samples must have ascending frames and complete RGB thumbnails")
        last_frame = sample[0]
    issues = []
    black_start = None
    white_start = None
    static_start = None
    previous = None
    previous_frame = None
    for frame, pixels in samples:
        if previous_frame is not None and frame - previous_frame > step_frames:
            # Missing observations must not connect independent runs or imply a static interval.
            if black_start is not None:
                issues.append({"code": "BLACK_FRAME", "severity": "error" if black_as_error else "warning",
                               "start_frame": black_start, "end_frame": previous_frame,
                               "message": "サンプリングしたプレビューがほぼ黒です（意図的な暗転の可能性あり）"})
                black_start = None
            if white_start is not None:
                issues.append({"code": "WHITE_FRAME", "severity": "error" if white_as_error else "warning",
                               "start_frame": white_start, "end_frame": previous_frame,
                               "message": "サンプリングしたプレビューがほぼ白です（意図的な白画面の可能性あり）"})
                white_start = None
            if static_start is not None and previous_frame - static_start >= min_static_frames:
                issues.append({"code": "STATIC_PREVIEW", "severity": "warning",
                               "start_frame": static_start, "end_frame": previous_frame,
                               "message": "プレビューの変化がほとんどありません（静止画の可能性あり）"})
            static_start = None
            previous = None
        black = (sum(max(pixels[i:i + 3]) < 12 for i in range(0, len(pixels), 3)) >=
                 len(pixels) / 3 * .99 and sum(pixels) / len(pixels) < 5)
        white = (sum(min(pixels[i:i + 3]) > 245 for i in range(0, len(pixels), 3)) >=
                 len(pixels) / 3 * .99 and sum(pixels) / len(pixels) > 250)
        if black and black_start is None:
            black_start = frame
        if not black and black_start is not None:
            issues.append({"code": "BLACK_FRAME", "severity": "error" if black_as_error else "warning",
                           "start_frame": black_start, "end_frame": previous_frame,
                           "message": "サンプリングしたプレビューがほぼ黒です（意図的な暗転の可能性あり）"})
            black_start = None
        if white and white_start is None:
            white_start = frame
        if not white and white_start is not None:
            issues.append({"code": "WHITE_FRAME", "severity": "error" if white_as_error else "warning",
                           "start_frame": white_start, "end_frame": previous_frame,
                           "message": "サンプリングしたプレビューがほぼ白です（意図的な白画面の可能性あり）"})
            white_start = None
        if previous is not None:
            unchanged = _visually_unchanged(previous, pixels)
            if unchanged and static_start is None:
                static_start = previous_frame
            if not unchanged and static_start is not None:
                if previous_frame - static_start >= min_static_frames:
                    issues.append({"code": "STATIC_PREVIEW", "severity": "warning",
                                   "start_frame": static_start, "end_frame": previous_frame,
                                   "message": "プレビューの変化がほとんどありません（静止画の可能性あり）"})
                static_start = None
        previous = pixels
        previous_frame = frame
    if samples:
        last = samples[-1][0]
        if black_start is not None:
            issues.append({"code": "BLACK_FRAME", "severity": "error" if black_as_error else "warning",
                           "start_frame": black_start, "end_frame": last,
                           "message": "サンプリングしたプレビューがほぼ黒です（意図的な暗転の可能性あり）"})
        if white_start is not None:
            issues.append({"code": "WHITE_FRAME", "severity": "error" if white_as_error else "warning",
                           "start_frame": white_start, "end_frame": last,
                           "message": "サンプリングしたプレビューがほぼ白です（意図的な白画面の可能性あり）"})
        if static_start is not None and last - static_start >= min_static_frames:
            issues.append({"code": "STATIC_PREVIEW", "severity": "warning",
                           "start_frame": static_start, "end_frame": last,
                           "message": "プレビューの変化がほとんどありません（静止画の可能性あり）"})
    return {"success": True, "passed": not any(i["severity"] == "error" for i in issues),
            "sample_count": len(samples), "issues": issues}
