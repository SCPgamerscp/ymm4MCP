"""Conservative preview QA on a bounded set of PNG captures."""

import base64
import binascii
import calendar
import zlib
from io import BytesIO

from PIL import Image, UnidentifiedImageError

from editing import MAX_FRAME
from png_scanlines import validate_idat_scanlines
from png_metadata_integrity import (validate_icc_profile, validate_language_tag,
                                    validate_text_payload)

# PNG sRGB requires the standardized gAMA/cHRM values when those chunks exist.
_SRGB_CHRM = b"".join(value.to_bytes(4, "big") for value in (
    31270, 32900, 64000, 33000, 30000, 60000, 15000, 6000))


def _inflate_png_metadata(data: bytes, limit: int = 2_000_000) -> bytes:
    """Decode ancillary zlib streams without accepting bombs or trailing bytes."""
    try:
        decoder = zlib.decompressobj()
        decoded = decoder.decompress(data, limit + 1)
        if (len(decoded) > limit or not decoder.eof or
                decoder.unused_data or decoder.unconsumed_tail):
            raise ValueError("invalid compressed PNG metadata")
        return decoded
    except zlib.error as exc:
        raise ValueError("invalid compressed PNG metadata") from exc


def _valid_png_keyword(keyword: bytes) -> bool:
    """PNG Latin-1 keyword: 1..79 printable bytes, with no edge/double spaces."""
    return (1 <= len(keyword) <= 79 and
            keyword[0] != 32 and keyword[-1] != 32 and
            b"  " not in keyword and
            all(32 <= byte <= 126 or 161 <= byte <= 255
                for byte in keyword))


def _validate_png_container(raw: bytes) -> None:
    """Reject malformed PNG chunks before Pillow discards their metadata.

    Validate critical PNG structure and color/palette semantics independently
    of Pillow, which may recover from malformed input while decoding pixels.
    """
    if not raw.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("invalid PNG signature")
    offset = 8
    chunk_count = 0
    seen_ihdr = seen_idat = seen_iend = idat_finished = False
    seen_plte = False
    color_type = bit_depth = None
    idat_payload_bytes = 0
    idat_chunks = []
    palette_entries = 0
    seen_trns = seen_bkgd = seen_sbit = seen_hist = False
    seen_gamma = seen_chrm = seen_srgb = seen_iccp = False
    seen_phys = seen_time = False
    gamma_value = chromaticity = None
    splt_names = set()
    # Bound the aggregate *expanded* metadata, not just each compressed chunk.
    # Many individually small zTXt/iTXt/iCCP chunks can otherwise expand to
    # gigabytes across a single (at most 10 MB) preview PNG.
    metadata_bytes = 0
    metadata_limit = 2_000_000
    while offset < len(raw):
        chunk_count += 1
        if chunk_count > 4096 or len(raw) - offset < 12:
            raise ValueError("invalid PNG chunk count or header")
        length = int.from_bytes(raw[offset:offset + 4], "big")
        kind = raw[offset + 4:offset + 8]
        end = offset + 12 + length
        if end > len(raw):
            raise ValueError("truncated PNG chunk")
        payload = raw[offset + 8:offset + 8 + length]
        expected_crc = int.from_bytes(raw[offset + 8 + length:end], "big")
        if zlib.crc32(raw[offset + 4:offset + 8 + length]) != expected_crc:
            raise ValueError("PNG chunk CRC mismatch")
        if not seen_ihdr:
            if kind != b"IHDR" or length != 13:
                raise ValueError("PNG must begin with IHDR")
            width = int.from_bytes(payload[:4], "big")
            height = int.from_bytes(payload[4:8], "big")
            bit_depth, color_type, compression, filtering, interlace = payload[8:]
            legal_depths = {
                0: (1, 2, 4, 8, 16),
                2: (8, 16),
                3: (1, 2, 4, 8),
                4: (8, 16),
                6: (8, 16),
            }
            if (not width or not height or width * height > 32_000_000 or
                    bit_depth not in legal_depths.get(color_type, ()) or
                    compression != 0 or filtering != 0 or interlace not in (0, 1)):
                raise ValueError("invalid PNG IHDR fields")
            seen_ihdr = True
        elif kind == b"IHDR":
            raise ValueError("duplicate PNG IHDR")
        if any(not (65 <= byte <= 90 or 97 <= byte <= 122) for byte in kind):
            raise ValueError("invalid PNG chunk type")
        if kind[2] & 0x20:
            raise ValueError("invalid PNG reserved chunk type")
        # Even an APNG with a single declared frame is not a static preview:
        # animation control changes which image the decoder may present.
        if kind in (b"acTL", b"fcTL", b"fdAT"):
            raise ValueError("animated PNG is not a static preview")
        if (kind[0] & 0x20) == 0 and kind not in (b"IHDR", b"PLTE", b"IDAT", b"IEND"):
            raise ValueError("unknown critical PNG chunk")
        if kind == b"PLTE":
            if seen_plte or seen_idat or seen_trns or seen_bkgd or seen_hist:
                raise ValueError("PNG PLTE must precede transparency and background metadata")
            if (color_type in (0, 4) or length == 0 or length > 768 or
                    length % 3 or
                    (color_type == 3 and length // 3 > 1 << bit_depth)):
                raise ValueError("invalid PNG palette")
            seen_plte = True
            palette_entries = length // 3
        if kind == b"sBIT":
            # Significant-bit metadata must precede the palette and image data.
            expected_channels = {0: 1, 2: 3, 3: 3, 4: 2, 6: 4}[color_type]
            maximum = 8 if color_type == 3 else bit_depth
            if (seen_sbit or seen_plte or seen_idat or
                    length != expected_channels or
                    any(value < 1 or value > maximum for value in payload)):
                raise ValueError("invalid PNG significant-bit metadata")
            seen_sbit = True
        if kind == b"tRNS":
            if seen_trns or seen_idat or color_type in (4, 6):
                raise ValueError("invalid PNG transparency metadata")
            if color_type == 3:
                if not seen_plte or not 1 <= length <= palette_entries:
                    raise ValueError("invalid indexed PNG transparency table")
            else:
                expected = 2 if color_type == 0 else 6
                if (length != expected or
                        any(int.from_bytes(payload[i:i + 2], "big") >= (1 << bit_depth)
                            for i in range(0, length, 2))):
                    raise ValueError("invalid PNG transparent color")
            seen_trns = True
        if kind == b"bKGD":
            if seen_bkgd or seen_idat:
                raise ValueError("invalid PNG background metadata")
            if color_type == 3:
                if not seen_plte or length != 1 or payload[0] >= palette_entries:
                    raise ValueError("invalid indexed PNG background")
            else:
                expected = 2 if color_type in (0, 4) else 6
                if (length != expected or
                        any(int.from_bytes(payload[i:i + 2], "big") >= (1 << bit_depth)
                            for i in range(0, length, 2))):
                    raise ValueError("invalid PNG background color")
            seen_bkgd = True
        if kind == b"hIST":
            if (seen_hist or seen_idat or color_type != 3 or
                    not seen_plte or length != 2 * palette_entries):
                raise ValueError("invalid PNG palette histogram")
            seen_hist = True
        if kind == b"gAMA":
            value = int.from_bytes(payload, "big")
            if (seen_gamma or seen_plte or seen_idat or length != 4 or
                    not value or (seen_srgb and value != 45455)):
                raise ValueError("invalid PNG gamma metadata")
            gamma_value = value
            seen_gamma = True
        if kind == b"cHRM":
            if (seen_chrm or seen_plte or seen_idat or length != 32 or
                    (seen_srgb and payload != _SRGB_CHRM)):
                raise ValueError("invalid PNG chromaticity metadata")
            chromaticity = payload
            seen_chrm = True
        if kind == b"sRGB":
            if (seen_srgb or seen_iccp or seen_plte or seen_idat or
                    length != 1 or payload[0] > 3 or
                    (gamma_value is not None and gamma_value != 45455) or
                    (chromaticity is not None and chromaticity != _SRGB_CHRM)):
                raise ValueError("invalid PNG sRGB metadata")
            seen_srgb = True
        if kind == b"iCCP":
            # Do not trust embedded profiles without checking the bounded
            # compressed payload; malformed ICC data may be ignored by Pillow.
            keyword, sep, remainder = payload.partition(b"\x00")
            if (seen_iccp or seen_srgb or seen_plte or seen_idat or
                    not sep or not _valid_png_keyword(keyword) or
                    len(remainder) < 2 or remainder[0] != 0):
                raise ValueError("invalid PNG ICC profile metadata")
            profile = _inflate_png_metadata(
                remainder[1:], limit=metadata_limit - metadata_bytes)
            validate_icc_profile(profile, color_type)
            metadata_bytes += len(profile)
            seen_iccp = True
        if kind == b"pHYs":
            if (seen_phys or seen_idat or length != 9 or
                    payload[8] not in (0, 1)):
                raise ValueError("invalid PNG physical pixel metadata")
            seen_phys = True
        if kind == b"tIME":
            if seen_time or length != 7:
                raise ValueError("invalid PNG modification time metadata")
            year = int.from_bytes(payload[:2], "big")
            month, day, hour, minute, second = payload[2:]
            if (not year or not 1 <= month <= 12 or
                    not 1 <= day <= calendar.monthrange(year, month)[1] or
                    hour > 23 or minute > 59 or second > 60):
                raise ValueError("invalid PNG modification time")
            seen_time = True
        if kind == b"sPLT":
            keyword, sep, remainder = payload.partition(b"\x00")
            if (seen_idat or not sep or not _valid_png_keyword(keyword) or
                    keyword in splt_names or not remainder or
                    remainder[0] not in (8, 16)):
                raise ValueError("invalid PNG suggested palette metadata")
            entry_size = 6 if remainder[0] == 8 else 10
            if len(remainder) <= 1 or (len(remainder) - 1) % entry_size:
                raise ValueError("invalid PNG suggested palette entries")
            splt_names.add(keyword)
        if kind == b"tEXt":
            keyword, sep, text = payload.partition(b"\x00")
            if (not sep or not _valid_png_keyword(keyword) or
                    len(text) > metadata_limit - metadata_bytes):
                raise ValueError("invalid PNG Latin-1 text metadata")
            validate_text_payload(text)
            metadata_bytes += len(text)
        if kind == b"zTXt":
            keyword, sep, remainder = payload.partition(b"\x00")
            if (not sep or not _valid_png_keyword(keyword) or
                    len(remainder) < 2 or remainder[0] != 0):
                raise ValueError("invalid PNG compressed text metadata")
            decoded_text = _inflate_png_metadata(
                remainder[1:], limit=metadata_limit - metadata_bytes)
            validate_text_payload(decoded_text)
            metadata_bytes += len(decoded_text)
        if kind == b"iTXt":
            keyword, sep, remainder = payload.partition(b"\x00")
            if (not sep or not _valid_png_keyword(keyword) or
                    len(remainder) < 4 or remainder[0] not in (0, 1) or
                    remainder[1] != 0):
                raise ValueError("invalid PNG international text header")
            language, sep, remainder_text = remainder[2:].partition(b"\x00")
            translated, translated_sep, encoded_text = remainder_text.partition(b"\x00")
            if not sep or not translated_sep:
                raise ValueError("invalid PNG international text language")
            validate_language_tag(language)
            validate_text_payload(translated)
            try:
                translated.decode("utf-8")
                decoded = (_inflate_png_metadata(
                    encoded_text, limit=metadata_limit - metadata_bytes)
                    if remainder[0] else encoded_text)
                if len(decoded) > metadata_limit - metadata_bytes:
                    raise ValueError("oversized PNG international text")
                decoded.decode("utf-8")
                validate_text_payload(decoded)
                metadata_bytes += len(decoded)
            except UnicodeDecodeError as exc:
                raise ValueError("invalid PNG international text UTF-8") from exc
        if kind == b"IDAT":
            if idat_finished:
                raise ValueError("nonconsecutive PNG IDAT chunks")
            if color_type == 3 and not seen_plte:
                raise ValueError("indexed PNG requires PLTE")
            seen_idat = True
            idat_payload_bytes += length
            idat_chunks.append(payload)
        elif seen_idat:
            idat_finished = True
        if kind == b"IEND":
            if length != 0 or not seen_idat or not idat_payload_bytes or end != len(raw):
                raise ValueError("invalid PNG end")
            seen_iend = True
            break
        offset = end
    if not seen_iend:
        raise ValueError("PNG is missing IEND")
    validate_idat_scanlines(idat_chunks, width, height, bit_depth, color_type, interlace)


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
