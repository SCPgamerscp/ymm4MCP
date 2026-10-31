"""Bounded structural validation of decompressed PNG image scanlines.

Pillow is deliberately not the authority for container validity: some decoders
recover pixels from a truncated stream or ignore trailing compressed data.
This module checks the exact scanline layout without materializing a whole
decompressed image in memory.
"""

import zlib


def _scanline_sizes(width: int, height: int, bits_per_pixel: int, interlace: int):
    if interlace == 0:
        yield from (1 + (width * bits_per_pixel + 7) // 8 for _ in range(height))
        return
    # PNG Adam7 pass origins and strides.
    for x0, y0, dx, dy in (
        (0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8),
        (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2),
        (0, 1, 1, 2),
    ):
        if width <= x0 or height <= y0:
            continue
        columns = (width - x0 + dx - 1) // dx
        rows = (height - y0 + dy - 1) // dy
        row_size = 1 + (columns * bits_per_pixel + 7) // 8
        yield from (row_size for _ in range(rows))


def validate_idat_scanlines(chunks: list[bytes], width: int, height: int,
                            bit_depth: int, color_type: int, interlace: int) -> None:
    """Require one complete zlib stream and exactly valid PNG filtered rows.

    The caller has already validated the PNG header, chunk CRCs, contiguous
    IDAT ordering, and a maximum pixel budget. Decompression output is
    processed in 64-KiB pieces so even 16-bit RGBA captures are bounded.
    """
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]
    sizes = iter(_scanline_sizes(width, height, channels * bit_depth, interlace))
    decoder = zlib.decompressobj()
    row_remaining = 0

    def consume(data: bytes) -> None:
        nonlocal row_remaining
        offset = 0
        while offset < len(data):
            if row_remaining == 0:
                row_remaining = next(sizes, 0)
                if row_remaining == 0:
                    raise ValueError("PNG has excess decompressed image data")
                if data[offset] > 4:
                    raise ValueError("invalid PNG scanline filter")
            take = min(row_remaining, len(data) - offset)
            offset += take
            row_remaining -= take

    try:
        for payload in chunks:
            if decoder.eof and payload:
                raise ValueError("PNG contains compressed data after zlib end")
            for start in range(0, len(payload), 65536):
                remaining_input = payload[start:start + 65536]
                while remaining_input:
                    previous_length = len(remaining_input)
                    decoded = decoder.decompress(remaining_input, 65536)
                    consume(decoded)
                    if decoder.unused_data:
                        raise ValueError("PNG has trailing compressed data")
                    remaining_input = decoder.unconsumed_tail
                    if decoder.eof:
                        if remaining_input:
                            raise ValueError("PNG has trailing compressed data")
                        break
                    if len(remaining_input) == previous_length and not decoded:
                        raise ValueError("PNG zlib decoder made no progress")
        if not decoder.eof:
            raise ValueError("truncated PNG compressed image stream")
        if row_remaining or next(sizes, None) is not None:
            raise ValueError("PNG decompressed scanline data is incomplete")
    except zlib.error as exc:
        raise ValueError("invalid PNG compressed image stream") from exc
