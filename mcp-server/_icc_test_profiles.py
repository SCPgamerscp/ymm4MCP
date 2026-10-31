"""Real ICC test fixtures for PNG preview validation (Issue #61).

The RGB profile comes from LittleCMS via Pillow; the gray profile is derived
from its real description, copyright, white point and tone curve tag payloads.
"""

from PIL import ImageCms


def rgb_icc_profile():
    return ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()


def gray_icc_profile():
    rgb = rgb_icc_profile()
    count = int.from_bytes(rgb[128:132], "big")
    tags = {}
    for i in range(count):
        entry = 132 + 12 * i
        signature = rgb[entry:entry + 4]
        offset = int.from_bytes(rgb[entry + 4:entry + 8], "big")
        size = int.from_bytes(rgb[entry + 8:entry + 12], "big")
        tags[signature] = rgb[offset:offset + size]

    chosen = ((b"desc", b"desc"), (b"cprt", b"cprt"),
              (b"wtpt", b"wtpt"), (b"kTRC", b"rTRC"))
    header = bytearray(rgb[:128])
    header[16:20] = b"GRAY"
    header[84:100] = bytes(16)  # Modified profiles have no stored profile ID.
    result = bytearray(header + len(chosen).to_bytes(4, "big") +
                       bytes(12 * len(chosen)))
    for i, (signature, source) in enumerate(chosen):
        while len(result) % 4:
            result.append(0)
        offset = len(result)
        payload = tags[source]
        result.extend(payload)
        entry = 132 + 12 * i
        result[entry:entry + 12] = (signature + offset.to_bytes(4, "big") +
                                    len(payload).to_bytes(4, "big"))
    while len(result) % 4:
        result.append(0)
    result[:4] = len(result).to_bytes(4, "big")
    return bytes(result)


def extended_rgb_icc_profile(size):
    """Add a well-formed optional characterization-target tag for budget tests."""
    rgb = rgb_icc_profile()
    if size < len(rgb) + 20 or size % 4:
        raise ValueError("ICC fixture size must be aligned and large enough")
    count = int.from_bytes(rgb[128:132], "big")
    result = bytearray(rgb[:128])
    result[84:100] = bytes(16)
    result.extend((count + 1).to_bytes(4, "big"))
    for i in range(count):
        entry = 132 + 12 * i
        offset = int.from_bytes(rgb[entry + 4:entry + 8], "big")
        result.extend(rgb[entry:entry + 4] +
                      (offset + 12).to_bytes(4, "big") +
                      rgb[entry + 8:entry + 12])
    payload_size = size - len(rgb) - 12
    result.extend(b"targ" + (len(rgb) + 12).to_bytes(4, "big") +
                  payload_size.to_bytes(4, "big"))
    result.extend(rgb[132 + 12 * count:])
    result.extend(b"text" + bytes(4) + b"D" * (payload_size - 8))
    result[:4] = len(result).to_bytes(4, "big")
    return bytes(result)
