"""Conservative preview QA on a bounded set of PNG captures."""

import base64
from io import BytesIO

from PIL import Image, UnidentifiedImageError


def thumbnail(image_b64: str) -> bytes:
    if not isinstance(image_b64, str) or len(image_b64) > 14_000_000:
        raise ValueError("preview image is missing or too large")
    try:
        raw = base64.b64decode(image_b64, validate=True)
        if len(raw) > 10_000_000:
            raise ValueError("preview image is too large")
        with Image.open(BytesIO(raw)) as image:
            if image.format != "PNG" or image.width * image.height > 32_000_000:
                raise ValueError("preview must be a PNG of at most 32 megapixels")
            image.load()
            return image.convert("RGB").resize((64, 36)).tobytes()
    except (ValueError, UnidentifiedImageError, OSError) as exc:
        raise ValueError("invalid preview PNG") from exc


def inspect(samples: list[tuple[int, bytes]], *, step_frames: int,
            min_static_frames: int = 60, black_as_error: bool = False) -> dict:
    """Only mark observed frames; gaps between samples are never called exact boundaries."""
    if not isinstance(samples, list) or len(samples) > 40:
        raise ValueError("samples must contain at most 40 captures")
    if isinstance(step_frames, bool) or not isinstance(step_frames, int) or step_frames < 1:
        raise ValueError("step_frames must be positive")
    if isinstance(min_static_frames, bool) or not isinstance(min_static_frames, int) or min_static_frames < 1:
        raise ValueError("min_static_frames must be positive")
    if not isinstance(black_as_error, bool):
        raise ValueError("black_as_error must be boolean")
    last_frame = -1
    for sample in samples:
        if (not isinstance(sample, tuple) or len(sample) != 2 or
                isinstance(sample[0], bool) or not isinstance(sample[0], int) or
                sample[0] <= last_frame or not isinstance(sample[1], bytes) or
                len(sample[1]) != 64 * 36 * 3):
            raise ValueError("samples must have ascending frames and complete RGB thumbnails")
        last_frame = sample[0]
    issues = []
    black_start = None
    static_start = None
    previous = None
    previous_frame = None
    for frame, pixels in samples:
        black = (sum(max(pixels[i:i + 3]) < 12 for i in range(0, len(pixels), 3)) >=
                 len(pixels) / 3 * .99 and sum(pixels) / len(pixels) < 5)
        if black and black_start is None:
            black_start = frame
        if not black and black_start is not None:
            issues.append({"code": "BLACK_FRAME", "severity": "error" if black_as_error else "warning",
                           "start_frame": black_start, "end_frame": previous_frame,
                           "message": "サンプリングしたプレビューがほぼ黒です（意図的な暗転の可能性あり）"})
            black_start = None
        if previous is not None:
            unchanged = sum(abs(a - b) for a, b in zip(previous, pixels)) / len(pixels) < 1
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
        if static_start is not None and last - static_start >= min_static_frames:
            issues.append({"code": "STATIC_PREVIEW", "severity": "warning",
                           "start_frame": static_start, "end_frame": last,
                           "message": "プレビューの変化がほとんどありません（静止画の可能性あり）"})
    return {"success": True, "passed": not any(i["severity"] == "error" for i in issues),
            "sample_count": len(samples), "issues": issues}
