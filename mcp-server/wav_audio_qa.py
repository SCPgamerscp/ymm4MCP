"""Bounded PCM WAV inspection for rendered-audio QA (Issue #61).

This is a standalone analyzer; it does not claim to inspect compressed MP4 audio.
"""
import math
import os
import struct
import wave


def inspect_wav(path, *, silence_db=-50.0, min_silence_seconds=1.0,
                clip_threshold=0.999, max_bytes=256_000_000):
    """Return deterministic clipping/silence findings without loading the whole file."""
    if not isinstance(path, (str, os.PathLike)):
        raise ValueError("path must be a filesystem path")
    if type(max_bytes) is not int or max_bytes <= 0:
        raise ValueError("max_bytes must be positive")
    if not all(isinstance(v, (float, int)) and not isinstance(v, bool)
               and math.isfinite(v) for v in (silence_db, min_silence_seconds, clip_threshold)):
        raise ValueError("thresholds must be finite numbers")
    if not -120 <= silence_db < 0 or min_silence_seconds <= 0 or not 0 < clip_threshold <= 1:
        raise ValueError("invalid thresholds")
    if os.stat(path).st_size > max_bytes:
        raise ValueError("WAV exceeds maximum input size")
    with wave.open(path, "rb") as wav:
        channels, width, rate, frames = (wav.getnchannels(), wav.getsampwidth(),
                                         wav.getframerate(), wav.getnframes())
        if channels < 1 or channels > 32 or width not in (1, 2, 3, 4) or rate < 1:
            raise ValueError("unsupported PCM WAV format")
        if frames * channels * width > max_bytes:
            raise ValueError("declared PCM payload exceeds maximum input size")
        threshold = 10 ** (silence_db / 20)
        silence_start = None
        silences = []
        peak = 0.0
        clipped = 0
        processed = 0
        while processed < frames:
            raw = wav.readframes(min(4096, frames - processed))
            frame_size = channels * width
            if len(raw) % frame_size:
                raise ValueError("truncated PCM frame")
            count = len(raw) // frame_size
            if count == 0:
                raise ValueError("truncated PCM WAV")
            for index in range(count):
                frame_peak = 0.0
                for ch in range(channels):
                    pos = (index * channels + ch) * width
                    sample = raw[pos:pos + width]
                    if width == 1:
                        value = (sample[0] - 128) / 128.0
                    else:
                        value = int.from_bytes(sample, "little", signed=True) / float(1 << (8 * width - 1))
                    magnitude = abs(value)
                    frame_peak = max(frame_peak, magnitude)
                    peak = max(peak, magnitude)
                    if magnitude >= clip_threshold:
                        clipped += 1
                at = processed + index
                if frame_peak <= threshold:
                    if silence_start is None:
                        silence_start = at
                elif silence_start is not None:
                    if (at - silence_start) / rate >= min_silence_seconds:
                        silences.append({"start_seconds": silence_start / rate, "end_seconds": at / rate})
                    silence_start = None
            processed += count
        if silence_start is not None and (frames - silence_start) / rate >= min_silence_seconds:
            silences.append({"start_seconds": silence_start / rate, "end_seconds": frames / rate})
        return {"duration_seconds": frames / rate, "sample_rate": rate,
                "channels": channels, "peak_dbfs": 20 * math.log10(peak) if peak else None,
                "clipped_samples": clipped, "long_silences": silences,
                "passed": clipped == 0 and not silences}
