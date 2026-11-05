"""Bounded PCM WAV inspection for rendered-audio QA (Issue #61).

This is a standalone analyzer; it does not claim to inspect compressed MP4 audio.
"""
import math
import os
import struct
import wave


def inspect_wav(path, *, silence_db=-50.0, min_silence_seconds=1.0,
                clip_threshold=0.999, max_bytes=256_000_000,
                max_stereo_imbalance_db=None, max_dc_offset_dbfs=None):
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
    if max_stereo_imbalance_db is not None and (
            type(max_stereo_imbalance_db) not in (int, float)
            or not math.isfinite(max_stereo_imbalance_db)
            or not 0 <= max_stereo_imbalance_db <= 120):
        raise ValueError("max_stereo_imbalance_db must be finite and in [0, 120]")
    if max_dc_offset_dbfs is not None and (
            type(max_dc_offset_dbfs) not in (int, float)
            or not math.isfinite(max_dc_offset_dbfs)
            or not -120 <= max_dc_offset_dbfs <= 0):
        raise ValueError("max_dc_offset_dbfs must be finite and in [-120, 0]")
    if os.stat(path).st_size > max_bytes:
        raise ValueError("WAV exceeds maximum input size")
    # wave.open on Python 3.12 treats pathlib.Path as a file object; pass a string.\n    with wave.open(os.fspath(path), "rb") as wav:
        channels, width, rate, frames = (wav.getnchannels(), wav.getsampwidth(),
                                         wav.getframerate(), wav.getnframes())
        if channels < 1 or channels > 32 or width not in (1, 2, 3, 4) or rate < 1:
            raise ValueError("unsupported PCM WAV format")
        if frames == 0:
            raise ValueError("WAV has no audio frames")
        if frames * channels * width > max_bytes:
            raise ValueError("declared PCM payload exceeds maximum input size")
        threshold = 10 ** (silence_db / 20)
        silence_start = None
        silences = []
        peak = 0.0
        clipped = 0
        channel_peaks = [0.0] * channels
        channel_sum_squares = [0.0] * channels
        channel_sums = [0.0] * channels
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
                    channel_peaks[ch] = max(channel_peaks[ch], magnitude)
                    channel_sum_squares[ch] += value * value
                    channel_sums[ch] += value
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
        silent_channels = [ch for ch, value in enumerate(channel_peaks) if value <= threshold]
        channel_rms = [math.sqrt(total / frames) for total in channel_sum_squares]
        rms = math.sqrt(sum(channel_sum_squares) / (frames * channels))
        # A stereo mix with one silent channel must never appear balanced.
        # Use RMS rather than sample peak to detect sustained channel imbalance.
        stereo_imbalance_db = None
        if channels == 2:
            if channel_rms[0] == 0 and channel_rms[1] == 0:
                stereo_imbalance_db = 0.0
            elif min(channel_rms) == 0:
                stereo_imbalance_db = None  # Infinite imbalance, JSON-safe.
            else:
                stereo_imbalance_db = 20 * math.log10(max(channel_rms) / min(channel_rms))
        imbalance_failed = (channels == 2 and max_stereo_imbalance_db is not None
                            and (stereo_imbalance_db is None
                                 or stereo_imbalance_db > max_stereo_imbalance_db))
        # DC offset is a non-zero mean waveform, not ordinary RMS loudness.
        channel_dc_offsets = [abs(total / frames) for total in channel_sums]
        dc_offset_dbfs = [20 * math.log10(value) if value else None
                          for value in channel_dc_offsets]
        dc_failed_channels = ([i for i, value in enumerate(channel_dc_offsets)
                               if value > 0 and
                               20 * math.log10(value) > max_dc_offset_dbfs]
                              if max_dc_offset_dbfs is not None else [])
        issues = []
        if dc_failed_channels:
            issues.append({"code": "AUDIO_DC_OFFSET",
                           "channels": dc_failed_channels,
                           "max_allowed_dbfs": max_dc_offset_dbfs})
        if clipped:
            issues.append({"code": "AUDIO_CLIPPING", "clipped_samples": clipped})
        for interval in silences:
            issues.append({"code": "AUDIO_LONG_SILENCE", **interval})
        if silent_channels:
            issues.append({"code": "AUDIO_SILENT_CHANNEL", "channels": silent_channels})
        if imbalance_failed:
            issues.append({"code": "AUDIO_STEREO_IMBALANCE",
                           "imbalance_db": stereo_imbalance_db,
                           "max_allowed_db": max_stereo_imbalance_db})
        return {"duration_seconds": frames / rate, "sample_rate": rate,
                "channels": channels, "peak_dbfs": 20 * math.log10(peak) if peak else None,
                "clipped_samples": clipped, "clipped_sample_fraction": clipped / (frames * channels),
                "long_silences": silences,
                "rms_dbfs": 20 * math.log10(rms) if rms else None,
                "channel_rms_dbfs": [20 * math.log10(value) if value else None
                                    for value in channel_rms],
                "channel_peak_dbfs": [20 * math.log10(value) if value else None
                                      for value in channel_peaks],
                "silent_channels": silent_channels,
                "channel_dc_offset_dbfs": dc_offset_dbfs,
                "stereo_imbalance_db": stereo_imbalance_db,
                "issues": issues,
                "passed": not issues}
