"""Strict, bounded validation for metadata in PNG preview evidence.

PNG decoders often ignore invalid ancillary metadata. For QA, a corrupted
color profile or ambiguous text record must not be treated as a valid capture.
"""


def validate_icc_profile(profile: bytes, color_type: int) -> None:
    """Check the ICC header and tag table without interpreting tag payloads."""
    if len(profile) < 132 or int.from_bytes(profile[:4], "big") != len(profile):
        raise ValueError("invalid PNG ICC profile size")
    if profile[36:40] != b"acsp":
        raise ValueError("invalid PNG ICC profile signature")
    expected_space = b"GRAY" if color_type in (0, 4) else b"RGB "
    if profile[16:20] != expected_space:
        raise ValueError("PNG ICC profile color space disagrees with IHDR")
    if profile[20:24] not in (b"XYZ ", b"Lab "):
        raise ValueError("invalid PNG ICC connection space")
    count = int.from_bytes(profile[128:132], "big")
    if count > (len(profile) - 132) // 12:
        raise ValueError("truncated PNG ICC tag table")
    for i in range(count):
        start = 132 + 12 * i
        offset = int.from_bytes(profile[start + 4:start + 8], "big")
        size = int.from_bytes(profile[start + 8:start + 12], "big")
        if (offset % 4 or size < 8 or offset < 132 + 12 * count or
                size > len(profile) - offset):
            raise ValueError("invalid PNG ICC tag bounds")


def validate_language_tag(tag: bytes) -> None:
    """Require nonempty, ASCII-alphanumeric language subtags."""
    if not tag:
        return  # PNG permits an unspecified language.
    parts = tag.split(b"-")
    if (not all(1 <= len(part) <= 8 for part in parts) or
            not all(65 <= ch <= 90 or 97 <= ch <= 122 for ch in parts[0]) or
            any(not all(65 <= ch <= 90 or 97 <= ch <= 122 or
                        48 <= ch <= 57 for ch in part) for part in parts)):
        raise ValueError("invalid PNG international text language tag")


def validate_text_payload(data: bytes) -> None:
    """Reject embedded NULs, which are forbidden in PNG text strings."""
    if b"\x00" in data:
        raise ValueError("PNG text contains embedded NUL")
