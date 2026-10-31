"""Strict, bounded validation for metadata in PNG preview evidence.

PNG decoders often ignore invalid ancillary metadata. For QA, a corrupted
color profile or ambiguous text record must not be treated as a valid capture.
"""


def validate_icc_profile(profile: bytes, color_type: int) -> None:
    """Validate mandatory ICC tags and bounded tag records without decoding colors."""
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
    if not count or count > (len(profile) - 132) // 12:
        raise ValueError("missing or truncated PNG ICC tag table")
    tags = {}
    for i in range(count):
        start = 132 + 12 * i
        signature = profile[start:start + 4]
        if signature in tags:
            raise ValueError("duplicate PNG ICC tag signature")
        offset = int.from_bytes(profile[start + 4:start + 8], "big")
        size = int.from_bytes(profile[start + 8:start + 12], "big")
        if (offset % 4 or size < 8 or offset < 132 + 12 * count or
                size > len(profile) - offset):
            raise ValueError("invalid PNG ICC tag bounds")
        tags[signature] = profile[offset:offset + size]

    # ICC.1 common required tags. A nonempty table alone is not sufficient:
    # a profile with arbitrary vendor tags still cannot describe its colors.
    required = {b"desc", b"cprt", b"wtpt"}
    if not required.issubset(tags):
        raise ValueError("missing required PNG ICC profile tags")
    tag_types = {b"desc": (b"desc", b"mluc"),
                 b"cprt": (b"text", b"mluc"),
                 b"wtpt": (b"XYZ ",)}
    if expected_space == b"GRAY":
        if b"kTRC" not in tags and b"A2B0" not in tags:
            raise ValueError("missing grayscale PNG ICC transform")
        if b"kTRC" in tags:
            tag_types[b"kTRC"] = (b"curv", b"para")
    else:
        matrix = {b"rXYZ", b"gXYZ", b"bXYZ",
                  b"rTRC", b"gTRC", b"bTRC"}
        if not matrix.issubset(tags) and b"A2B0" not in tags:
            raise ValueError("missing RGB PNG ICC color transform")
        if profile[12:16] == b"mntr" and not matrix.issubset(tags) and b"B2A0" not in tags:
            raise ValueError("incomplete RGB display PNG ICC LUT transform")
        for signature in matrix & tags.keys():
            tag_types[signature] = ((b"XYZ ",) if signature.endswith(b"XYZ")
                                    else (b"curv", b"para"))
    for signature, allowed in tag_types.items():
        payload = tags[signature]
        if len(payload) < 12 or payload[:4] not in allowed or payload[4:8] != bytes(4):
            raise ValueError("invalid required PNG ICC tag payload")
        if signature.endswith(b"XYZ") or signature == b"wtpt":
            if len(payload) < 20:
                raise ValueError("truncated PNG ICC XYZ tag")


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
