"""Read-only, bounded FFmpeg decode QA for completed local video exports."""
import json
import math
import re
import subprocess
import threading
import time
from pathlib import Path

from jobs import is_absolute_media_path

VIDEO_EXTENSIONS = {".mp4", ".mov", ".mkv", ".webm", ".avi"}


def _unique_object_pairs(pairs):
    """Reject duplicate ffprobe JSON keys instead of trusting the last value."""
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate ffprobe metadata key")
        result[key] = value
    return result


def _signal_ranges(log, start_name, end_name, window_duration=None):
    """Parse FFmpeg filter event pairs; incomplete events are not measurements."""
    pattern = re.compile(r"\b(" + re.escape(start_name) + r"|" + re.escape(end_name)
                         + r"):\s*([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)(?=$|[\s,;])")
    start = None
    ranges = []
    for line in log.splitlines():
        for name, timestamp in pattern.findall(line):
            value = float(timestamp)
            if not math.isfinite(value):
                # Do not serialize Infinity or create an invented signal span.
                start = None
                continue
            # FFmpeg may report negative silence_start during filter preroll.
            # Clamp it rather than turning '-1.5' into a positive 1.5.
            value = max(0.0, value)
            if window_duration is not None:
                value = min(value, window_duration)
            if name == start_name:
                start = value
            elif start is not None:
                if value > start:
                    ranges.append((start, value))
                start = None
    # A black/silent interval that reaches EOF has no matching *_end line.
    # Close it at the decoded window boundary rather than silently dropping it.
    if start is not None and window_duration is not None and start < window_duration:
        ranges.append((start, window_duration))
    return ranges


def _audio_peak_db(log):
    """Highest finite sample peak from FFmpeg astats (dBFS); not true-peak."""
    peaks = []
    for value in re.findall(r"\bPeak level dB:\s*([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)(?=$|[\s,;])", log):
        peak = float(value)
        if math.isfinite(peak):
            peaks.append(peak)
    return max(peaks) if peaks else None


def _optional_stream_time(value, *, signed=False):
    """Parse optional ffprobe stream times without accepting ambiguous values."""
    if value is None or value == "N/A":
        return None
    sign = r"[+-]?" if signed else ""
    if (type(value) is not str or not re.fullmatch(
            sign + r"(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?",
            value)):
        raise ValueError("invalid stream time")
    seconds = float(value)
    if not math.isfinite(seconds) or (not signed and seconds < 0):
        raise ValueError("invalid stream time")
    return seconds


def _frame_rate(value):
    """Parse ffprobe's exact rational FPS; unknown 0/x is not measured FPS."""
    if value is None or value == "N/A":
        return None
    if type(value) is not str or not re.fullmatch(r"[0-9]{1,9}/[0-9]{1,9}", value):
        raise ValueError("invalid frame rate")
    numerator, denominator = map(int, value.split("/"))
    if denominator == 0:
        if numerator == 0:
            return None
        raise ValueError("invalid frame rate denominator")
    if numerator == 0:
        return None
    fps = numerator / denominator
    if not 0.1 <= fps <= 1000:
        raise ValueError("frame rate outside supported range")
    return round(fps, 3)


def _reject_nonfinite_json(value):
    """Strict JSON does not permit NaN or Infinity tokens."""
    raise ValueError("non-standard JSON number: " + value)


def _check_probe_depth(value, max_depth=64):
    """Bound nested ffprobe metadata, including unused fields, without recursion."""
    pending = [(value, 0)]
    while pending:
        node, depth = pending.pop()
        if depth > max_depth:
            raise ValueError("ffprobe metadata nesting limit exceeded")
        if isinstance(node, dict):
            pending.extend((child, depth + 1) for child in node.values())
        elif isinstance(node, list):
            pending.extend((child, depth + 1) for child in node)


def _file_identity(st):
    """Include change time so restoring mtime cannot mask an in-place edit."""
    return (st.st_dev, st.st_ino, st.st_size, st.st_mtime_ns,
            getattr(st, "st_ctime_ns", None))


def failure(code, message):
    return {"success": False, "passed": False, "issues": [], "error_code": code, "error": message}


def run(command, timeout):
    """Capture bounded subprocess diagnostics without exhausting RAM or disk.

    Both pipes are drained concurrently. A child exceeding either output cap
    is killed and rejected, rather than silently treating truncated output
    as successful media verification.
    """
    stdout_limit = 1_000_001  # One byte beyond the accepted ffprobe JSON limit.
    stderr_limit = 1_000_000
    stdout_head = bytearray()
    stderr_tail = bytearray()
    errors = []
    over_limit = threading.Event()

    with subprocess.Popen(command, stdin=subprocess.DEVNULL,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE) as process:
        def drain(pipe, keep_head):
            total = 0
            limit = stdout_limit if keep_head else stderr_limit
            try:
                while True:
                    chunk = pipe.read(65_536)
                    if not chunk:
                        break
                    total += len(chunk)
                    if keep_head:
                        room = stdout_limit - len(stdout_head)
                        if room > 0:
                            stdout_head.extend(chunk[:room])
                    else:
                        stderr_tail.extend(chunk)
                        if len(stderr_tail) > stderr_limit:
                            del stderr_tail[:-stderr_limit]
                    if total > limit:
                        over_limit.set()
                        # Stop the noisy process immediately. The other reader
                        # keeps draining until the child closes both pipes.
                        try:
                            process.kill()
                        except ProcessLookupError:
                            pass
                        break
            except OSError as exc:
                errors.append(exc)
            finally:
                pipe.close()

        stdout_thread = threading.Thread(target=drain, args=(process.stdout, True),
                                         daemon=True)
        stderr_thread = threading.Thread(target=drain, args=(process.stderr, False),
                                         daemon=True)
        stdout_thread.start()
        stderr_thread.start()
        try:
            returncode = process.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
            stdout_thread.join()
            stderr_thread.join()
            raise
        stdout_thread.join()
        stderr_thread.join()
    if errors:
        raise OSError("failed to capture FFmpeg process output") from errors[0]
    if over_limit.is_set():
        raise OSError("FFmpeg subprocess output exceeded the capture limit")
    return subprocess.CompletedProcess(
        command, returncode, stdout_head.decode("utf-8"),
        stderr_tail.decode("utf-8", errors="replace"))

def inspect(path, *, require_audio=True, sample_seconds=8, full_decode=False):
    """Sample both ends and optionally decode all video/audio streams end-to-end."""
    if not is_absolute_media_path(path) or Path(path).suffix.lower() not in VIDEO_EXTENSIONS:
        raise ValueError("path must be an absolute video file path")
    if not isinstance(require_audio, bool):
        raise ValueError("require_audio must be boolean")
    if type(sample_seconds) is not int or not 1 <= sample_seconds <= 15:
        raise ValueError("sample_seconds must be an integer in 1..15")
    if not isinstance(full_decode, bool):
        raise ValueError("full_decode must be boolean")
    target = Path(path)
    try:
        exists = target.is_file()
    except OSError:
        return failure("EXPORT_VERIFY_FAILED", "動画ファイルを読み取れません")
    if not exists:
        return failure("EXPORT_FILE_MISSING", "動画ファイルがありません（Python MCP側のローカルパスを指定）")
    try:
        initial_stat = target.stat()
        if initial_stat.st_size < 32:
            return failure("EXPORT_FILE_EMPTY", "動画ファイルが小さすぎます")
    except OSError:
        return failure("EXPORT_VERIFY_FAILED", "動画ファイルを読み取れません")

    try:
        probe = run(["ffprobe", "-v", "error", "-show_entries",
                     "format=duration,start_time:stream=index,codec_type,width,height,avg_frame_rate,sample_rate,channels,duration,start_time",
                     "-of", "json", str(target)], 15)
    except FileNotFoundError:
        return failure("FFMPEG_UNAVAILABLE", "ffprobeが見つかりません")
    except subprocess.TimeoutExpired:
        return failure("EXPORT_PROBE_TIMEOUT", "ffprobeが時間切れです")
    except OSError:
        return failure("EXPORT_PROBE_FAILED", "ffprobeを実行できません")
    except UnicodeDecodeError:
        return failure("EXPORT_PROBE_FAILED", "ffprobeの出力が不正なUTF-8です")
    if probe.returncode or len(probe.stdout.encode("utf-8")) > 1_000_000:
        return failure("EXPORT_PROBE_FAILED", "ffprobeで動画を解析できません")
    try:
        info = json.loads(probe.stdout, object_pairs_hook=_unique_object_pairs,
                          parse_constant=_reject_nonfinite_json)
        _check_probe_depth(info)
        streams = info["streams"]
        raw_duration = info["format"]["duration"]
        # Python float() accepts underscores, Unicode digits, and whitespace,
        # none of which is an unambiguous ffprobe decimal duration.
        if (type(raw_duration) not in (int, float, str)
                or (isinstance(raw_duration, str) and not re.fullmatch(
                    r"(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?",
                    raw_duration))):
            raise ValueError("invalid duration syntax")
        duration = float(raw_duration)
        # Transport streams can start at a nonzero presentation timestamp.
        # Compare track positions to the container origin, not wall-clock zero.
        container_start = _optional_stream_time(
            info["format"].get("start_time"), signed=True)
        # When format.start_time is absent, derive the origin from validated
        # A/V stream timestamps below instead of assuming wall-clock zero.
        # The FFmpeg command uses microsecond precision; smaller durations
        # would round -t to zero and could falsely pass an empty sample.
        if not isinstance(streams, list) or not math.isfinite(duration) or duration < 0.000001:
            raise ValueError("invalid metadata")
        # The A/V cap below does not cover subtitle/data/attachment tracks.
        # Bound the entire stream table before iterating over untrusted metadata.
        if len(streams) > 128:
            return failure("EXPORT_STREAM_LIMIT", "ストリーム数が多すぎます（最大128）")
        seen_indices = set()
        for stream in streams:
            if (not isinstance(stream, dict)
                    or not isinstance(stream.get("codec_type"), str)
                    or stream["codec_type"] not in (
                        "video", "audio", "subtitle", "data", "attachment", "unknown")):
                raise ValueError("invalid stream metadata")
            if "index" in stream:
                index = stream["index"]
                if (type(index) is not int or index < 0 or index > 2_147_483_647
                        or index in seen_indices):
                    raise ValueError("invalid or duplicate stream index")
                seen_indices.add(index)
            # ffprobe may expose video dimensions even when the codec cannot
            # actually be decoded. Incomplete or impossible dimensions must
            # not be reported as a verified video stream.
            if stream["codec_type"] == "video" and (
                    "width" in stream or "height" in stream):
                width, height = stream.get("width"), stream.get("height")
                if (type(width) is not int or type(height) is not int
                        or not 1 <= width <= 16_384 or not 1 <= height <= 16_384):
                    raise ValueError("invalid video dimensions")
            if stream["codec_type"] == "video" and "avg_frame_rate" in stream:
                _frame_rate(stream["avg_frame_rate"])
            # The presence of an audio stream is not evidence of a usable
            # audio format when ffprobe reports zero or malformed parameters.
            if stream["codec_type"] in ("video", "audio"):
                if "duration" in stream:
                    _optional_stream_time(stream["duration"])
                if "start_time" in stream:
                    _optional_stream_time(stream["start_time"], signed=True)
            if stream["codec_type"] == "audio":
                if "sample_rate" in stream:
                    rate = stream["sample_rate"]
                    if (not isinstance(rate, str) or not rate.isascii()
                            or not rate.isdecimal() or not 1 <= int(rate) <= 768_000):
                        raise ValueError("invalid audio sample rate")
                if "channels" in stream:
                    channels = stream["channels"]
                    if type(channels) is not int or not 1 <= channels <= 64:
                        raise ValueError("invalid audio channel count")
        if container_start is None:
            known_starts = [
                _optional_stream_time(stream.get("start_time"), signed=True)
                for stream in streams if stream["codec_type"] in ("video", "audio")]
            # A missing timestamp on any A/V track makes the origin ambiguous.
            # Do not shift the timeline using only a late-starting audio track.
            container_start = (min(known_starts) if known_starts
                               and all(value is not None for value in known_starts)
                               else 0.0)
        for stream in streams:
            if stream["codec_type"] in ("video", "audio"):
                stream_start = _optional_stream_time(stream.get("start_time"), signed=True)
                if stream_start is not None and stream_start - container_start > duration + 1:
                    raise ValueError("stream starts beyond container end")
    except (ValueError, TypeError, KeyError, OverflowError, RecursionError):
        return failure("EXPORT_PROBE_FAILED", "動画の尺・ストリーム情報が不正です")

    # Huge or pathological durations lose the precision needed to sample the
    # tail reliably. Refuse to certify a video whose end cannot be checked.
    if duration > 7 * 24 * 60 * 60:
        return failure("EXPORT_DURATION_LIMIT", "動画の長さが検証上限（7日）を超えています")

    # Full decode is intentionally bounded: it is opt-in, not a background job.
    # Reject oversized work rather than silently claiming a full verification.
    if full_decode and duration > 900:
        return failure("EXPORT_FULL_DECODE_DURATION_LIMIT",
                       "全尺デコードは15分以内の動画に限定しています")

    videos = [s for s in streams if isinstance(s, dict) and s.get("codec_type") == "video"]
    audios = [s for s in streams if isinstance(s, dict) and s.get("codec_type") == "audio"]
    # Sample decoding is synchronous: cap total stream work before launching
    # per-stream subprocesses, including streams not required by the caller.
    if len(videos) + len(audios) > 16:
        return failure("EXPORT_STREAM_LIMIT", "映像・音声ストリームが多すぎます（最大16）")
    video_details = [
        {"stream_ordinal": i, "width": stream.get("width"),
         "height": stream.get("height"),
         "fps": _frame_rate(stream.get("avg_frame_rate"))}
        for i, stream in enumerate(videos)]
    audio_details = [
        {"stream_ordinal": i,
         "sample_rate": int(stream["sample_rate"]) if "sample_rate" in stream else None,
         "channels": stream.get("channels")}
        for i, stream in enumerate(audios)]
    issues = []
    # Multiple output tracks may intentionally differ, so surface format
    # inconsistencies as QA warnings rather than rejecting the export.
    dimensions = {f'{v["width"]}x{v["height"]}' for v in video_details
                  if v["width"] is not None}
    frame_rates = {v["fps"] for v in video_details if v["fps"] is not None}
    sample_rates = {a["sample_rate"] for a in audio_details
                    if a["sample_rate"] is not None}
    channel_counts = {a["channels"] for a in audio_details
                      if a["channels"] is not None}
    for values, code, field in (
            (dimensions, "EXPORT_VIDEO_TRACK_RESOLUTION_MISMATCH", "resolutions"),
            (frame_rates, "EXPORT_VIDEO_TRACK_FPS_MISMATCH", "frame_rates"),
            (sample_rates, "EXPORT_AUDIO_TRACK_SAMPLE_RATE_MISMATCH", "sample_rates"),
            (channel_counts, "EXPORT_AUDIO_TRACK_CHANNEL_MISMATCH", "channel_counts")):
        if len(values) > 1:
            issues.append({"code": code, "severity": "warning",
                           field: sorted(values)})
    # A successful tail decode only proves that some samples exist in the
    # final window. Track metadata can reveal an earlier end within it.
    for kind, tracks in (("video", videos), ("audio", audios)):
        for ordinal, stream in enumerate(tracks):
            track_duration = _optional_stream_time(stream.get("duration"))
            track_start = _optional_stream_time(stream.get("start_time"), signed=True)
            if track_start is None:
                track_start = container_start
            track_start -= container_start
            if track_start is not None and track_start > 1.0:
                issues.append({"code": "EXPORT_STREAM_STARTS_LATE",
                               "severity": "warning", "stream_type": kind,
                               "stream_ordinal": ordinal,
                               "missing_head_seconds": round(track_start, 3)})
            if track_start < -1.0:
                # Negative preroll is allowed, but excessive offsets should
                # be visible to the unattended editor as a QA warning.
                issues.append({"code": "EXPORT_STREAM_STARTS_BEFORE_CONTAINER",
                               "severity": "warning", "stream_type": kind,
                               "stream_ordinal": ordinal,
                               "preroll_seconds": round(-track_start, 3)})
            if track_duration is not None and track_start is not None:
                if track_start + track_duration > duration + 1.0:
                    issues.append({"code": "EXPORT_STREAM_EXCEEDS_CONTAINER",
                                   "severity": "warning", "stream_type": kind,
                                   "stream_ordinal": ordinal,
                                   "excess_seconds": round(
                                       track_start + track_duration - duration, 3)})
                missing = duration - (track_start + track_duration)
                if missing > 1.0:
                    issues.append({"code": "EXPORT_STREAM_ENDS_EARLY",
                                   "severity": "warning", "stream_type": kind,
                                   "stream_ordinal": ordinal,
                                   "missing_tail_seconds": round(missing, 3)})
    if not videos:
        issues.append({"code": "EXPORT_VIDEO_STREAM_MISSING", "severity": "error"})
    if require_audio and not audios:
        issues.append({"code": "EXPORT_AUDIO_STREAM_MISSING", "severity": "error"})
    starts = [0.0]
    # Cover the middle of longer exports, not only the beginning and tail.
    # Three disjoint windows are enough to catch a broken middle segment
    # without turning synchronous sampling into an unbounded full decode.
    if duration >= 3 * sample_seconds:
        starts.append((duration - sample_seconds) / 2)
    if duration > sample_seconds:
        starts.append(max(0.0, duration - sample_seconds))
    samples = []
    black_frames = []
    frozen_frames = []
    long_silences = []
    audio_peaks = []
    # Per-command timeouts alone allow up to 32 slow subprocesses. Bound
    # the complete synchronous sampling pass to two minutes.
    sample_deadline = time.monotonic() + 120
    for start in starts:
        length = min(float(sample_seconds), duration - start)
        if length <= 0:
            continue
        samples.append({"start_seconds": round(start, 6), "duration_seconds": round(length, 6)})
        # A successful first track cannot certify another broken track.
        for kind, ordinal, stream in (
                [("video", i, track) for i, track in enumerate(videos)]
                + [("audio", i, track) for i, track in enumerate(audios)]):
            # Keep the decode check and add non-blocking signal hints. The
            # filters emit bounded info logs for the sampled windows only.
            signal_filter = ("blackdetect=d=1:pix_th=0.10:pic_th=0.98,"
                              "freezedetect=n=-60dB:d=2"
                             if kind == "video" else "silencedetect=noise=-50dB:d=2,astats=metadata=0:reset=0")
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-nostats",
                       "-v", "info", "-xerror", "-ss", format(start, ".6f"),
                       "-i", str(target), "-t", format(length, ".6f"),
                       "-map", f"0:v:{ordinal}" if kind == "video" else f"0:a:{ordinal}",
                       "-vf" if kind == "video" else "-af", signal_filter,
                       "-abort_on", "empty_output", "-f", "null", "-"]
            remaining = sample_deadline - time.monotonic()
            if remaining <= 0:
                return failure("EXPORT_DECODE_TIMEOUT", "動画サンプルの総検証時間を超えました")
            try:
                decoded = run(command, min(30, remaining))
            except FileNotFoundError:
                return failure("FFMPEG_UNAVAILABLE", "ffmpegが見つかりません")
            except subprocess.TimeoutExpired:
                return failure("EXPORT_DECODE_TIMEOUT", "動画サンプルのデコードが時間切れです")
            except OSError:
                return failure("EXPORT_DECODE_FAILED", "ffmpegを実行できません")
            if decoded.returncode:
                issues.append({"code": "EXPORT_DECODE_FAILED", "severity": "error",
                               "stream_type": kind, "stream_ordinal": ordinal,
                               "start_seconds": round(start, 6),
                               "details": (decoded.stderr or "")[-500:]})
                continue
            if kind == "audio":
                peak = _audio_peak_db(decoded.stderr or "")
                if peak is not None:
                    peak_event = {"sample_start_seconds": round(start, 6),
                                  "stream_ordinal": ordinal,
                                  "peak_dbfs": round(peak, 3)}
                    audio_peaks.append(peak_event)
                    # Near-0 dBFS samples may be clipped, but a peak alone
                    # cannot prove distortion or inter-sample true-peak.
                    if peak >= -0.1:
                        issues.append({"code": "EXPORT_SAMPLE_AUDIO_PEAK_NEAR_CLIPPING",
                                       "severity": "warning", **peak_event})
            # Black/silent stretches can be intentional (e.g. a fade). Report
            # them as warnings, not as proof of an invalid export.
            start_name, end_name = (("black_start", "black_end") if kind == "video"
                                    else ("silence_start", "silence_end"))
            for signal_start, signal_end in _signal_ranges(
                    decoded.stderr or "", start_name, end_name, length):
                event = {"sample_start_seconds": round(start, 6),
                         "stream_ordinal": ordinal,
                         "filter_start_seconds": round(signal_start, 3),
                         "filter_end_seconds": round(signal_end, 3),
                          "video_start_seconds": round(start + signal_start, 6),
                          "video_end_seconds": round(start + signal_end, 6),
                         "duration_seconds": round(signal_end - signal_start, 3)}
                if kind == "video":
                    black_frames.append(event)
                    code = "EXPORT_SAMPLE_BLACK_FRAMES"
                else:
                    long_silences.append(event)
                    code = "EXPORT_SAMPLE_LONG_SILENCE"
                issues.append({"code": code, "severity": "warning", **event})
            if kind == "video":
                for freeze_start, freeze_end in _signal_ranges(
                        decoded.stderr or "", "freeze_start", "freeze_end", length):
                    event = {"sample_start_seconds": round(start, 6),
                             "stream_ordinal": ordinal,
                             "filter_start_seconds": round(freeze_start, 3),
                             "filter_end_seconds": round(freeze_end, 3),
                              "video_start_seconds": round(start + freeze_start, 6),
                              "video_end_seconds": round(start + freeze_end, 6),
                             "duration_seconds": round(freeze_end - freeze_start, 3)}
                    frozen_frames.append(event)
                    issues.append({"code": "EXPORT_SAMPLE_FROZEN_FRAMES",
                                   "severity": "warning", **event})
    full_decode_ok = False
    if full_decode and videos:
        full_decode_ok = True
        # -map 0:v / 0:a checks every video/audio stream, not only stream 0.
        # No signal filters are used here: sample-only diagnostics above remain
        # bounded, and full-file stderr is restricted to decoding errors.
        timeout = min(900, max(60, int(math.ceil(duration * 4))))
        # Both video and audio share a single full-decode budget.
        full_deadline = time.monotonic() + 900
        for kind, present in (("video", bool(videos)), ("audio", bool(audios))):
            if not present:
                continue
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-nostats",
                       "-v", "error", "-xerror", "-i", str(target),
                       "-map", "0:v" if kind == "video" else "0:a",
                       # Every mapped stream must produce decoded output. A
                       # populated first track must not hide an empty second.
                       "-abort_on", "empty_output+empty_output_stream",
                       "-f", "null", "-"]
            remaining = full_deadline - time.monotonic()
            if remaining <= 0:
                full_decode_ok = False
                issues.append({"code": "EXPORT_FULL_DECODE_TIMEOUT",
                               "severity": "error", "stream_type": kind})
                break
            try:
                decoded = run(command, min(timeout, remaining))
            except FileNotFoundError:
                return failure("FFMPEG_UNAVAILABLE", "ffmpegが見つかりません")
            except subprocess.TimeoutExpired:
                full_decode_ok = False
                issues.append({"code": "EXPORT_FULL_DECODE_TIMEOUT",
                               "severity": "error", "stream_type": kind})
                continue
            except OSError:
                full_decode_ok = False
                issues.append({"code": "EXPORT_FULL_DECODE_FAILED",
                               "severity": "error", "stream_type": kind})
                continue
            if decoded.returncode:
                full_decode_ok = False
                issues.append({"code": "EXPORT_FULL_DECODE_FAILED",
                               "severity": "error", "stream_type": kind,
                               "details": (decoded.stderr or "")[-500:]})
    # A video replaced or modified during ffprobe/ffmpeg cannot be certified
    # using observations from different file versions.
    try:
        final_stat = target.stat()
    except OSError:
        return failure("EXPORT_VERIFY_FAILED", "検証後に動画ファイルを読み取れません")
    if _file_identity(initial_stat) != _file_identity(final_stat):
        return failure("EXPORT_FILE_CHANGED", "検証中に動画ファイルが変更されました")
    passed = not any(issue["severity"] == "error" for issue in issues)
    fully_decoded = bool(full_decode and full_decode_ok and passed)
    # A requested full decode is not proof it completed. Do not label a
    # failed or interrupted verification as full_audio_video_decode.
    scope = ("full_audio_video_decode" if fully_decoded
             else "incomplete" if not passed
             else "sampled_audio_video_decode")
    return {"success": True,
            "passed": passed,
            "issues": issues,
            "verification_scope": scope,
            "fully_decoded": fully_decoded,
            "signal_scope": "sampled_only", "black_frames": black_frames,
            "frozen_frames": frozen_frames,
            "video_details": video_details, "audio_details": audio_details,
            "resolution": (f'{video_details[0]["width"]}x{video_details[0]["height"]}'
                           if video_details and video_details[0]["width"] is not None
                           else None),
            "fps": video_details[0]["fps"] if video_details else None,
            "long_silences": long_silences, "audio_peaks": audio_peaks,
            "duration_seconds": duration, "video_streams": len(videos),
            "audio_streams": len(audios), "samples": samples, "path": str(target)}
