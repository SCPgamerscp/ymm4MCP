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
            return image.convert("RGB").resize((64, 36)).convert("L").tobytes()
    except (ValueError, UnidentifiedImageError, OSError) as exc:
        raise ValueError("invalid preview PNG") from exc


def inspect(samples: list[tuple[int, bytes]], *, step_frames: int,
            min_static_frames: int = 60, black_as_error: bool = False) -> dict:
    """Only mark observed frames; gaps between samples are never called exact boundaries."""
    issues = []
    black_start = None
    static_start = None
    previous = None
    for frame, pixels in samples:
        black = sum(p < 12 for p in pixels) >= len(pixels) * .99 and sum(pixels) / len(pixels) < 5
        if black and black_start is None:
            black_start = frame
        if not black and black_start is not None:
            issues.append({"code": "BLACK_FRAME", "severity": "error" if black_as_error else "warning",
                           "start_frame": black_start, "end_frame": frame - step_frames,
                           "message": "サンプリングしたプレビューがほぼ黒です（意図的な暗転の可能性あり）"})
            black_start = None
        if previous is not None:
            unchanged = sum(abs(a - b) for a, b in zip(previous, pixels)) / len(pixels) < 1
            if unchanged and static_start is None:
                static_start = frame - step_frames
            if not unchanged and static_start is not None:
                if frame - step_frames - static_start >= min_static_frames:
                    issues.append({"code": "STATIC_PREVIEW", "severity": "warning",
                                   "start_frame": static_start, "end_frame": frame - step_frames,
                                   "message": "プレビューの変化がほとんどありません（静止画の可能性あり）"})
                static_start = None
        previous = pixels
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
