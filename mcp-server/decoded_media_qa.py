"""Read-only, bounded FFmpeg decode QA for completed local video exports."""
import json
import math
import re
import subprocess
import threading
import time
from fractions import Fraction
from pathlib import Path

from jobs import is_absolute_media_path

VIDEO_EXTENSIONS = {".mp4", ".mov", ".mkv", ".webm", ".avi"}
MAX_SIGNAL_EVENTS = 256  # Per signal type, across all tracks and windows.


def _unique_object_pairs(pairs):
    """Reject duplicate ffprobe JSON keys instead of trusting the last value."""
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate ffprobe metadata key")
        result[key] = value
    return result


def _trusted_filter_payload(line, filter_name, bare_names):
    """Ignore media metadata and unrelated FFmpeg logs masquerading as QA."""
    line = line.lstrip()
    # Tests and callers may provide the filter's unprefixed diagnostic text.
    if any(line.startswith(name + ":") for name in bare_names):
        return line
    match = re.match(
        r"^\[(?:Parsed_)?" + re.escape(filter_name)
        + r"(?:_[0-9]+)?\s+@\s+[^\]]+\]\s*(.*)$", line)
    return match.group(1) if match else ""


def _signal_ranges(log, start_name, end_name, window_duration=None):
    """Parse FFmpeg filter event pairs; incomplete events are not measurements."""
    pattern = re.compile(r"\b(" + re.escape(start_name) + r"|" + re.escape(end_name)
                         + r"):\s*([+-]?(?:(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?|nan|inf(?:inity)?))(?=$|[\s,;])", re.IGNORECASE)
    start = None
    ranges = []
    filter_name = {"black_start": "blackdetect", "silence_start": "silencedetect",
                   "freeze_start": "freezedetect"}[start_name]
    # Bare diagnostics are supported for older FFmpeg logs and fixtures,
    # but another filter\'s line must never supply this filter\'s events.
    bare_names = (start_name, end_name)
    for line in log.splitlines():
        payload = _trusted_filter_payload(line, filter_name, bare_names)
        for name, timestamp in pattern.findall(payload):
            # A malformed non-finite end invalidates the active interval.
            # Otherwise an unmatched start would be extended to sample EOF
            # and invent a black/silent/frozen span.
            value = float(timestamp)
            if not math.isfinite(value):
                start = None
                continue
            # FFmpeg may report negative silence_start during filter preroll.
            # Clamp it rather than turning '-1.5' into a positive 1.5.
            value = max(0.0, value)
            if window_duration is not None:
                value = min(value, window_duration)
            if name == start_name:
                # Repeated starts in concatenated/overlapping diagnostics
                # must not discard the beginning of a real signal interval.
                start = value if start is None else min(start, value)
            elif start is not None and value > start:
                ranges.append((start, value))
                start = None
            # An out-of-order end is not evidence that the active signal
            # stopped; retain its start for a later valid end.
    # A black/silent interval that reaches EOF has no matching *_end line.
    # Close it at the decoded window boundary rather than silently dropping it.
    if start is not None and window_duration is not None and start < window_duration:
        ranges.append((start, window_duration))
    return ranges


def _audio_peak_db(log):
    """Highest finite sample peak from FFmpeg astats (dBFS); not true-peak."""
    peaks = []
    pattern = re.compile(r"\bPeak level dB:\s*([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)(?=$|[\s,;])")
    for line in log.splitlines():
        payload = _trusted_filter_payload(line, "astats", ("Peak level dB",))
        for value in pattern.findall(payload):
            peak = float(value)
            if math.isfinite(peak):
                peaks.append(peak)
    return max(peaks) if peaks else None



def _audio_channel_peaks(log):
    """Read channel-specific FFmpeg astats peaks, including -inf for silence.

    A separate Overall section is never a channel measurement. Only astats
    diagnostics are trusted; media title/comment text must not supply peaks.
    """
    peaks = {}
    channel = None
    # FFmpeg channel counts are capped at 64. Never convert arbitrarily
    # long untrusted decimal fields to Python integers.
    channel_pattern = re.compile(r"Channel:\s*([1-9][0-9]{0,2})")
    peak_pattern = re.compile(
        r"Peak level dB:\s*(-inf|[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)",
        re.IGNORECASE)
    for line in log.splitlines():
        payload = _trusted_filter_payload(
            line, "astats", ("Channel", "Overall", "Peak level dB")).strip()
        if payload == "Overall":
            channel = None
            continue
        if payload.startswith("Channel:"):
            # Invalid channel headers must also clear the previous channel,
            # otherwise the following peak can be attributed to the wrong one.
            channel = None
            match = channel_pattern.fullmatch(payload)
            if match:
                value = int(match.group(1))
                if value <= 64:
                    channel = value
            continue
        if channel is None:
            continue
        match = peak_pattern.fullmatch(payload)
        if match:
            raw = match.group(1)
            peak = -math.inf if raw.lower() == "-inf" else float(raw)
            if peak == -math.inf or math.isfinite(peak):
                # Multiple astats summaries may occur in one decode. The
                # last one can be silent even when earlier audio was loud;
                # retain the highest measured peak across the whole sample.
                peaks[channel] = max(peaks.get(channel, -math.inf), peak)
    return peaks



def _audio_overall_silent(log):
    """Detect an all-silent astats Overall peak even without channel sections.

    Some FFmpeg builds emit only the Overall summary. Trust only astats
    diagnostics, and use the last reported Overall peak if repeated.
    """
    in_overall = False
    silent = None
    for line in log.splitlines():
        payload = _trusted_filter_payload(
            line, "astats", ("Channel", "Overall", "Peak level dB")).strip()
        if payload == "Overall":
            in_overall = True
        elif payload.startswith("Channel:"):
            in_overall = False
        elif in_overall and payload.startswith("Peak level dB:"):
            # Overall summaries can repeat. A later silent summary must
            # not erase evidence of an earlier non-silent segment.
            silent = (silent is not False and re.fullmatch(
                r"Peak level dB:\s*-inf", payload, re.IGNORECASE) is not None)
    return silent is True


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


def _frame_rate_ratio(value):
    """Expose exact normalized FPS alongside the rounded display number."""
    if _frame_rate(value) is None:
        return None
    numerator, denominator = map(int, value.split("/"))
    ratio = Fraction(numerator, denominator)
    return f"{ratio.numerator}/{ratio.denominator}"


def _reject_nonfinite_json(value):
    """Strict JSON does not permit NaN or Infinity tokens."""
    raise ValueError("non-standard JSON number: " + value)


def _finite_json_float(value):
    """Reject JSON exponent overflow, even in otherwise unused metadata."""
    number = float(value)
    if not math.isfinite(number):
        raise ValueError("non-finite JSON number")
    return number


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

def inspect(path, *, require_audio=True, sample_seconds=8, full_decode=False,
            expected_duration_seconds=None, duration_tolerance_seconds=1.0):
    """Sample completed media, optionally checking its expected final duration."""
    if not is_absolute_media_path(path) or Path(path).suffix.lower() not in VIDEO_EXTENSIONS:
        raise ValueError("path must be an absolute video file path")
    if not isinstance(require_audio, bool):
        raise ValueError("require_audio must be boolean")
    if type(sample_seconds) is not int or not 1 <= sample_seconds <= 15:
        raise ValueError("sample_seconds must be an integer in 1..15")
    if not isinstance(full_decode, bool):
        raise ValueError("full_decode must be boolean")
    if (expected_duration_seconds is not None and
            (type(expected_duration_seconds) not in (int, float) or
             not math.isfinite(expected_duration_seconds) or
             not 0 < expected_duration_seconds <= 7 * 24 * 60 * 60)):
        raise ValueError("expected_duration_seconds must be finite and in (0, 604800]")
    if (type(duration_tolerance_seconds) not in (int, float) or
            not math.isfinite(duration_tolerance_seconds) or
            not 0 <= duration_tolerance_seconds <= 60):
        raise ValueError("duration_tolerance_seconds must be finite and in [0, 60]")
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
                     "format=duration,start_time,nb_streams,size:stream=index,codec_type,width,height,avg_frame_rate,sample_rate,channels,duration,start_time,nb_frames,time_base:stream_disposition=attached_pic",
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
                          parse_constant=_reject_nonfinite_json,
                          parse_float=_finite_json_float)
        _check_probe_depth(info)
        streams = info["streams"]
        # An independent ffprobe size claim must match the file we inspected.
        # A contradictory or malformed claim is not a trustworthy export.
        reported_size = info["format"].get("size")
        if reported_size is not None:
            if (type(reported_size) is not str
                    or not re.fullmatch(r"[0-9]{1,20}", reported_size)
                    or int(reported_size) != initial_stat.st_size):
                raise ValueError("ffprobe file size disagrees with disk")
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
        # ffprobe reports the total stream count independently of the stream
        # array. An incomplete or contradictory inventory must not be certified.
        declared_streams = info["format"].get("nb_streams")
        if declared_streams is not None and (
                type(declared_streams) is not int or declared_streams != len(streams)):
            raise ValueError("ffprobe stream count disagrees with stream table")
        seen_indices = set()
        last_index = -1
        for stream in streams:
            if (not isinstance(stream, dict)
                    or not isinstance(stream.get("codec_type"), str)
                    or stream["codec_type"] not in (
                        "video", "audio", "subtitle", "data", "attachment", "unknown")):
                raise ValueError("invalid stream metadata")
            # Cover artwork is a video stream to ffprobe, but not a timed
            # video track. Do not count or decode it as ordinary video.
            disposition = stream.get("disposition", {})
            if not isinstance(disposition, dict):
                raise ValueError("invalid stream disposition")
            attached_pic = disposition.get("attached_pic", 0)
            if (type(attached_pic) is not int or attached_pic not in (0, 1)
                    or (attached_pic and stream["codec_type"] != "video")):
                raise ValueError("invalid attached-picture disposition")
            if "index" in stream:
                index = stream["index"]
                if (type(index) is not int or index < 0 or index > 2_147_483_647
                        or index in seen_indices):
                    raise ValueError("invalid or duplicate stream index")
                # FFmpeg's 0:v:N/0:a:N selectors follow stream index order.
                # Reject reordered metadata so diagnostics cannot be attributed
                # to the wrong stream when per-type ordinals are constructed.
                if index <= last_index:
                    raise ValueError("out-of-order stream indices")
                last_index = index
                seen_indices.add(index)
            if (stream["codec_type"] in ("video", "audio") and not attached_pic
                    and "time_base" in stream):
                # A zero or malformed timestamp unit makes all stream timing
                # comparisons unreliable, even if duration happens to parse.
                time_base = stream["time_base"]
                if (type(time_base) is not str or not re.fullmatch(
                        r"[0-9]{1,18}/[0-9]{1,18}", time_base)):
                    raise ValueError("invalid stream time base")
                unit, scale = map(int, time_base.split("/"))
                if unit == 0 or scale == 0:
                    raise ValueError("zero stream time base")
            if (stream["codec_type"] == "video" and not attached_pic
                    and "nb_frames" in stream):
                frame_count = stream["nb_frames"]
                # ffprobe uses N/A when the frame count is unknown. A reported
                # zero or malformed count must not be mistaken for a valid track.
                if frame_count not in (None, "N/A"):
                    if (type(frame_count) is not str
                            or not re.fullmatch(r"[0-9]{1,19}", frame_count)
                            or not 1 <= int(frame_count) <= 9_223_372_036_854_775_807):
                        raise ValueError("invalid video frame count")
            # ffprobe may expose video dimensions even when the codec cannot
            # actually be decoded. Incomplete or impossible dimensions must
            # not be reported as a verified video stream.
            if (stream["codec_type"] == "video" and not attached_pic and (
                    "width" in stream or "height" in stream)):
                width, height = stream.get("width"), stream.get("height")
                if (type(width) is not int or type(height) is not int
                        or not 1 <= width <= 16_384 or not 1 <= height <= 16_384
                        # A 16K-by-16K frame may allocate >1 GiB per frame.
                        # Keep decoded pixel buffers within a bounded 64 Mi-pixel budget.
                        or width * height > 67_108_864):
                    raise ValueError("invalid video dimensions")
            if (stream["codec_type"] == "video" and not attached_pic
                    and "avg_frame_rate" in stream):
                _frame_rate(stream["avg_frame_rate"])
            # The presence of an audio stream is not evidence of a usable
            # audio format when ffprobe reports zero or malformed parameters.
            if stream["codec_type"] in ("video", "audio") and not attached_pic:
                if "duration" in stream:
                    stream_duration = _optional_stream_time(stream["duration"])
                    # An explicitly empty A/V track cannot be certified by
                    # decoding some other stream or a permissive FFmpeg build.
                    if stream_duration == 0:
                        raise ValueError("zero-duration audio/video stream")
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
        # Real ffprobe stream tables either provide an index for every
        # stream or for none. A partial table cannot be reconciled reliably
        # with FFmpeg's per-type stream selectors.
        if seen_indices and len(seen_indices) != len(streams):
            raise ValueError("partially indexed stream table")
        # Absolute selectors are needed to skip cover artwork. Legacy mocks
        # without indices remain supported only when there is no cover art.
        has_cover_art = any(s.get("disposition", {}).get("attached_pic") == 1
                            for s in streams)
        if has_cover_art and len(seen_indices) != len(streams):
            raise ValueError("cover-art stream indices are unavailable")
        if container_start is None:
            known_starts = [
                _optional_stream_time(stream.get("start_time"), signed=True)
                for stream in streams if stream["codec_type"] in ("video", "audio")
                and stream.get("disposition", {}).get("attached_pic") != 1]
            # A missing timestamp on any A/V track makes the origin ambiguous.
            # Do not shift the timeline using only a late-starting audio track.
            container_start = (min(known_starts) if known_starts
                               and all(value is not None for value in known_starts)
                               else 0.0)
        for stream in streams:
            if (stream["codec_type"] in ("video", "audio")
                    and stream.get("disposition", {}).get("attached_pic") != 1):
                stream_start = _optional_stream_time(stream.get("start_time"), signed=True)
                if stream_start is not None:
                    relative_start = stream_start - container_start
                    if not math.isfinite(relative_start):
                        raise ValueError("stream start offset overflow")
                    if relative_start > duration + 1:
                        raise ValueError("stream starts beyond container end")
                    # Finite individual timestamps can still overflow when
                    # added. Reject them before QA emits Infinity in JSON.
                    track_length = _optional_stream_time(stream.get("duration"))
                    if track_length is not None and not math.isfinite(
                            relative_start + track_length):
                        raise ValueError("stream end offset overflow")
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

    videos = [s for s in streams if s["codec_type"] == "video"
              and s.get("disposition", {}).get("attached_pic") != 1]
    audios = [s for s in streams if isinstance(s, dict) and s.get("codec_type") == "audio"]
    # Sample decoding is synchronous: cap total stream work before launching
    # per-stream subprocesses, including streams not required by the caller.
    if len(videos) + len(audios) > 16:
        return failure("EXPORT_STREAM_LIMIT", "映像・音声ストリームが多すぎます（最大16）")
    video_details = [
        {"stream_ordinal": i, "stream_index": stream.get("index"),
         "width": stream.get("width"), "height": stream.get("height"),
         "fps": _frame_rate(stream.get("avg_frame_rate")),
         "fps_ratio": _frame_rate_ratio(stream.get("avg_frame_rate")),
         "reported_frames": (int(stream["nb_frames"])
                             if stream.get("nb_frames") not in (None, "N/A")
                             else None)}
        for i, stream in enumerate(videos)]
    audio_details = [
        {"stream_ordinal": i, "stream_index": stream.get("index"),
         "sample_rate": int(stream["sample_rate"]) if "sample_rate" in stream else None,
         "channels": stream.get("channels")}
        for i, stream in enumerate(audios)]
    issues = []
    # An exported file can decode correctly while being truncated or longer
    # than the edit plan. This is an error only when the caller supplied an
    # explicit expected duration, with a bounded tolerance.
    if (expected_duration_seconds is not None and
            abs(duration - expected_duration_seconds) > duration_tolerance_seconds):
        issues.append({
            "code": "EXPORT_DURATION_MISMATCH", "severity": "error",
            "expected_duration_seconds": expected_duration_seconds,
            "actual_duration_seconds": duration,
            "difference_seconds": round(duration - expected_duration_seconds, 6),
            "tolerance_seconds": duration_tolerance_seconds})
    # Multiple output tracks may intentionally differ, so surface format
    # inconsistencies as QA warnings rather than rejecting the export.
    dimensions = {f'{v["width"]}x{v["height"]}' for v in video_details
                  if v["width"] is not None}
    # Keep exact rational frame rates for comparisons. The public fps field
    # is rounded for display, and distinct rates can round to the same value.
    exact_frame_rates = set()
    for stream in videos:
        raw_rate = stream.get("avg_frame_rate")
        if raw_rate not in (None, "N/A") and _frame_rate(raw_rate) is not None:
            numerator, denominator = map(int, raw_rate.split("/"))
            exact_frame_rates.add(Fraction(numerator, denominator))
    sample_rates = {a["sample_rate"] for a in audio_details
                    if a["sample_rate"] is not None}
    channel_counts = {a["channels"] for a in audio_details
                      if a["channels"] is not None}
    for values, code, field in (
            (dimensions, "EXPORT_VIDEO_TRACK_RESOLUTION_MISMATCH", "resolutions"),
            (exact_frame_rates, "EXPORT_VIDEO_TRACK_FPS_MISMATCH", "frame_rates"),
            (sample_rates, "EXPORT_AUDIO_TRACK_SAMPLE_RATE_MISMATCH", "sample_rates"),
            (channel_counts, "EXPORT_AUDIO_TRACK_CHANNEL_MISMATCH", "channel_counts")):
        if len(values) > 1:
            if field == "frame_rates":
                # Preserve both human-readable approximations and the exact
                # ratios, so e.g. 30000/1001 and 2997/100 remain distinguishable.
                issues.append({"code": code, "severity": "warning",
                               "frame_rates": sorted({round(float(v), 3) for v in values}),
                               "frame_rate_ratios": sorted(
                                   f"{v.numerator}/{v.denominator}" for v in values)})
            else:
                issues.append({"code": code, "severity": "warning",
                               field: sorted(values)})
    # A successful tail decode only proves that some samples exist in the
    # final window. Track metadata can reveal an earlier end within it.
    for kind, tracks in (("video", videos), ("audio", audios)):
        for ordinal, stream in enumerate(tracks):
            track_duration = _optional_stream_time(stream.get("duration"))
            track_start = _optional_stream_time(stream.get("start_time"), signed=True)
            # An unknown track origin is not evidence that the track starts
            # at the container origin. Avoid invented end/gap warnings.
            if track_start is not None:
                track_start -= container_start
            if track_start is not None and track_start > 1.0:
                issues.append({"code": "EXPORT_STREAM_STARTS_LATE",
                               "severity": "warning", "stream_type": kind,
                               "stream_ordinal": ordinal,
                               "missing_head_seconds": round(track_start, 3)})
            if track_start is not None and track_start < -1.0:
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
    # A brief secondary track can fall entirely between the three regular
    # windows. When its validated timing is known, add one bounded targeted
    # window rather than falsely certifying a track that was never decoded.
    # At most one additional window per A/V stream is added (maximum 16).
    for stream in videos + audios:
        track_start = _optional_stream_time(stream.get("start_time"), signed=True)
        track_duration = _optional_stream_time(stream.get("duration"))
        if track_start is None or track_duration is None:
            continue
        left = max(0.0, track_start - container_start)
        right = min(duration, track_start - container_start + track_duration)
        if right <= left:
            continue
        # A few milliseconds of overlap at a window boundary do not
        # meaningfully sample a short track. Require at least 100 ms of
        # overlap, or half the track for tracks shorter than 200 ms.
        required_overlap = min(0.1, (right - left) / 2)
        if any(min(right, window + sample_seconds) - max(left, window)
               >= required_overlap for window in starts):
            continue
        # Place the target within the actual track; leave enough room for
        # FFmpeg's microsecond timestamp rounding and short preroll.
        center = (left + right) / 2
        window = max(0.0, min(center - sample_seconds / 2,
                              max(0.0, duration - sample_seconds)))
        starts.append(window)
    starts.sort()
    samples = []
    black_frames = []
    frozen_frames = []
    long_silences = []
    audio_peaks = []
    audio_channel_peaks = []
    # Overlapping decode windows can clip one physical defect into different
    # absolute ranges. Coalesce overlapping intervals per signal type/track
    # and keep the public event and its separate QA issue in sync.
    signal_records = {"video": [], "audio": [], "freeze": []}
    # A track outside all chosen windows must not be silently certified.
    sampled_tracks = set()

    def record_signal(kind, event, code, destination):
        records = signal_records[kind]
        start = event["video_start_seconds"]
        end = event["video_end_seconds"]
        matches = []
        # A new interval can bridge two previously disjoint detections.
        # Expand until the union is stable (transitive overlap).
        while True:
            additional = [(stored, issue) for stored, issue in records
                          if stored["stream_ordinal"] == event["stream_ordinal"]
                          and not any(stored is prior for prior, _ in matches)
                          and stored["video_start_seconds"] <= end
                          and start <= stored["video_end_seconds"]]
            if not additional:
                break
            matches.extend(additional)
            start = min(start, *(stored["video_start_seconds"]
                                 for stored, _ in additional))
            end = max(end, *(stored["video_end_seconds"]
                             for stored, _ in additional))
        if not matches:
            # Distinct findings are scanned during overlap coalescing. Bound
            # the count to avoid quadratic work on pathological filter logs.
            if len(records) >= MAX_SIGNAL_EVENTS:
                return False
            issue = {"code": code, "severity": "warning", **event}
            destination.append(event)
            issues.append(issue)
            records.append((event, issue))
            return True
        anchor_event, anchor_issue = matches[0]
        for redundant_event, redundant_issue in matches[1:]:
            records[:] = [(stored, issue) for stored, issue in records
                          if stored is not redundant_event]
            destination[:] = [stored for stored in destination
                              if stored is not redundant_event]
            issues[:] = [issue for issue in issues
                         if issue is not redundant_issue]
        # sample_start/filter_* describe the first observation; the video_*
        # coordinates and duration describe the merged absolute union.
        for target in (anchor_event, anchor_issue):
            target["video_start_seconds"] = round(start, 6)
            target["video_end_seconds"] = round(end, 6)
            target["duration_seconds"] = round(end - start, 3)
        return True
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
            # A late-starting or early-ending track can have no frames in
            # a sample window. Skip only with validated ffprobe timestamps;
            # absence of a track in this window is not a decode failure.
            track_start = _optional_stream_time(stream.get("start_time"), signed=True)
            track_duration = _optional_stream_time(stream.get("duration"))
            if track_start is not None:
                relative_start = track_start - container_start
                if start + length <= relative_start or (
                        track_duration is not None
                        and start >= relative_start + track_duration):
                    continue
            # A stream can end inside the requested sample window. FFmpeg
            # then emits no *_end event, but the missing media after its end
            # must not be counted as observed black/silent/frozen content.
            signal_window_length = length
            if track_start is not None and track_duration is not None:
                signal_window_length = min(
                    length, max(0.0, relative_start + track_duration - start))
            # Stream timestamps are absolute, but FFmpeg filter timestamps
            # are relative to this sample. Never report preroll before the
            # known stream start as black/silent/frozen media.
            signal_active_start = (max(0.0, relative_start - start)
                                   if track_start is not None else 0.0)
            # Keep the decode check and add non-blocking signal hints. The
            # filters emit bounded info logs for the sampled windows only.
            signal_filter = ("blackdetect=d=1:pix_th=0.10:pic_th=0.98,"
                              "freezedetect=n=-60dB:d=2"
                             if kind == "video" else "silencedetect=noise=-50dB:d=2,astats=metadata=0:reset=0")
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-nostats",
                       "-v", "info", "-xerror", "-ss", format(start, ".6f"),
                       "-i", str(target), "-t", format(length, ".6f"),
                       "-map", (f"0:{stream['index']}" if kind == "video" and has_cover_art
                                else f"0:v:{ordinal}" if kind == "video"
                                else f"0:a:{ordinal}"),
                       "-vf" if kind == "video" else "-af", signal_filter,
                       "-abort_on", "empty_output+empty_output_stream", "-f", "null", "-"]
            remaining = sample_deadline - time.monotonic()
            if remaining <= 0:
                return failure("EXPORT_DECODE_TIMEOUT", "動画サンプルの総検証時間を超えました")
            try:
                decoded = run(command, min(30, remaining))
            except FileNotFoundError:
                return failure("FFMPEG_UNAVAILABLE", "ffmpegが見つかりません")
            except subprocess.TimeoutExpired:
                return failure("EXPORT_DECODE_TIMEOUT", "動画サンプルのデコードが時間切れです")
            except (OSError, UnicodeDecodeError):
                return failure("EXPORT_DECODE_FAILED", "ffmpegの出力を読み取れません")
            if decoded.returncode:
                issues.append({"code": "EXPORT_DECODE_FAILED", "severity": "error",
                               "stream_type": kind, "stream_ordinal": ordinal,
                               "start_seconds": round(start, 6),
                               "details": (decoded.stderr or "")[-500:]})
                continue
            sampled_tracks.add((kind, ordinal))
            if kind == "audio":
                channel_peaks = _audio_channel_peaks(decoded.stderr or "")
                declared_channels = audio_details[ordinal]["channels"]
                if channel_peaks:
                    # JSON has no representation for -Infinity. Preserve a
                    # separate silent-channel list instead of emitting NaN/Inf.
                    audio_channel_peaks.append({
                        "sample_start_seconds": round(start, 6),
                        "stream_ordinal": ordinal,
                        "peak_dbfs_by_channel": {
                            str(channel): (round(value, 3)
                                           if math.isfinite(value) else None)
                            for channel, value in sorted(channel_peaks.items())},
                        "silent_channels": sorted(
                            channel for channel, value in channel_peaks.items()
                            if value == -math.inf)})
                overall_silent = _audio_overall_silent(decoded.stderr or "")
                complete_channels = (channel_peaks and declared_channels is not None
                                     and set(channel_peaks) ==
                                     set(range(1, declared_channels + 1)))
                if channel_peaks and declared_channels is not None:
                    expected = set(range(1, declared_channels + 1))
                    observed = set(channel_peaks)
                    if observed != expected:
                        issues.append({
                            "code": "EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH",
                            "severity": "warning", "stream_ordinal": ordinal,
                            "sample_start_seconds": round(start, 6),
                            "declared_channels": declared_channels,
                            "observed_channels": sorted(observed)})
                    elif declared_channels == 2:
                        left, right = channel_peaks[1], channel_peaks[2]
                        if ((math.isfinite(left) and math.isfinite(right)
                             and max(left, right) > -60
                              and abs(left - right) >= 12)
                                or (left == -math.inf and right > -60)
                                or (right == -math.inf and left > -60)):
                            issues.append({
                                "code": "EXPORT_SAMPLE_AUDIO_STEREO_IMBALANCE",
                                "severity": "warning", "stream_ordinal": ordinal,
                                "sample_start_seconds": round(start, 6),
                                "left_peak_dbfs": round(left, 3) if math.isfinite(left) else None,
                                "right_peak_dbfs": round(right, 3) if math.isfinite(right) else None})
                # A complete channel inventory or an Overall -inf peak
                # suggests silence, but neither can override a finite peak
                # measured elsewhere in this same sample. Inconsistent or
                # repeated astats sections must not produce a false all-silent
                # finding when any channel/Overall actually contained sound.
                peak = _audio_peak_db(decoded.stderr or "")
                if (peak is None and (overall_silent or (complete_channels and all(
                        value == -math.inf for value in channel_peaks.values())))):
                    issues.append({
                        "code": "EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                        "severity": "warning", "stream_ordinal": ordinal,
                        "sample_start_seconds": round(start, 6)})
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
                    decoded.stderr or "", start_name, end_name, signal_window_length):
                signal_start = max(signal_start, signal_active_start)
                if signal_start >= signal_end:
                    continue
                event = {"sample_start_seconds": round(start, 6),
                         "stream_ordinal": ordinal,
                         "filter_start_seconds": round(signal_start, 3),
                         "filter_end_seconds": round(signal_end, 3),
                          "video_start_seconds": round(start + signal_start, 6),
                          "video_end_seconds": round(start + signal_end, 6),
                         "duration_seconds": round(signal_end - signal_start, 3)}
                if kind == "video":
                    if not record_signal("video", event, "EXPORT_SAMPLE_BLACK_FRAMES",
                                         black_frames):
                        return failure("EXPORT_SIGNAL_LIMIT",
                                       "黒画面検出イベント数が上限を超えました")
                else:
                    if not record_signal("audio", event, "EXPORT_SAMPLE_LONG_SILENCE",
                                         long_silences):
                        return failure("EXPORT_SIGNAL_LIMIT",
                                       "無音検出イベント数が上限を超えました")
            if kind == "video":
                for freeze_start, freeze_end in _signal_ranges(
                        decoded.stderr or "", "freeze_start", "freeze_end",
                        signal_window_length):
                    freeze_start = max(freeze_start, signal_active_start)
                    if freeze_start >= freeze_end:
                        continue
                    event = {"sample_start_seconds": round(start, 6),
                             "stream_ordinal": ordinal,
                             "filter_start_seconds": round(freeze_start, 3),
                             "filter_end_seconds": round(freeze_end, 3),
                              "video_start_seconds": round(start + freeze_start, 6),
                              "video_end_seconds": round(start + freeze_end, 6),
                             "duration_seconds": round(freeze_end - freeze_start, 3)}
                    if not record_signal("freeze", event, "EXPORT_SAMPLE_FROZEN_FRAMES",
                                         frozen_frames):
                        return failure("EXPORT_SIGNAL_LIMIT",
                                       "フリーズ検出イベント数が上限を超えました")
    # A short track may lie entirely between sample windows. Do not claim
    # that QA passed when that track was never actually decoded.
    for kind, tracks in (("video", videos), ("audio", audios)):
        for ordinal in range(len(tracks)):
            if (kind, ordinal) not in sampled_tracks:
                issues.append({"code": "EXPORT_STREAM_NOT_SAMPLED",
                               "severity": "error", "stream_type": kind,
                               "stream_ordinal": ordinal})
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
            mapping_args = ["-map", "0:v" if kind == "video" else "0:a"]
            if kind == "video" and has_cover_art:
                mapping_args = [part for stream in videos
                                for part in ("-map", f"0:{stream['index']}")]
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-nostats",
                       "-v", "error", "-xerror", "-i", str(target),
                       *mapping_args,
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
            except (OSError, UnicodeDecodeError):
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
    scope = ("full_audio_video_decode" if fully_decoded and audios
             else "full_video_decode" if fully_decoded
             else "incomplete" if not passed
             else "sampled_audio_video_decode" if audios
             else "sampled_video_decode")
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
            # Highest observed sample peak across every audio track/window.
            # None means no finite measurement, not proof of silence.
            "audio_peak_db": max((p["peak_dbfs"] for p in audio_peaks), default=None),
            "audio_channel_peaks": audio_channel_peaks,
            "duration_seconds": duration, "video_streams": len(videos),
            "audio_streams": len(audios), "samples": samples, "path": str(target)}
