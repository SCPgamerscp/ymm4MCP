"""Read-only, bounded FFmpeg decode QA for completed local video exports."""
import json
import math
import re
import subprocess
from pathlib import Path

from jobs import is_absolute_media_path

VIDEO_EXTENSIONS = {".mp4", ".mov", ".mkv", ".webm", ".avi"}


def _signal_ranges(log, start_name, end_name):
    """Parse FFmpeg filter event pairs; incomplete events are not measurements."""
    pattern = re.compile(r"\b(" + re.escape(start_name) + r"|" + re.escape(end_name)
                         + r"):\s*(\d+(?:\.\d+)?)")
    start = None
    ranges = []
    for line in log.splitlines():
        for name, timestamp in pattern.findall(line):
            value = float(timestamp)
            if name == start_name:
                start = value
            elif start is not None:
                if value > start:
                    ranges.append((start, value))
                start = None
    return ranges


def _audio_peak_db(log):
    """Highest finite sample peak from FFmpeg astats (dBFS); not true-peak."""
    peaks = []
    for value in re.findall(r"\bPeak level dB:\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\b", log):
        peak = float(value)
        if math.isfinite(peak):
            peaks.append(peak)
    return max(peaks) if peaks else None


def failure(code, message):
    return {"success": False, "passed": False, "issues": [], "error_code": code, "error": message}


def run(command, timeout):
    return subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE, text=True, timeout=timeout, check=False)


def inspect(path, *, require_audio=True, sample_seconds=8):
    """Decode up to 15s at the start and end; do not claim full-file verification."""
    if not is_absolute_media_path(path) or Path(path).suffix.lower() not in VIDEO_EXTENSIONS:
        raise ValueError("path must be an absolute video file path")
    if not isinstance(require_audio, bool):
        raise ValueError("require_audio must be boolean")
    if type(sample_seconds) is not int or not 1 <= sample_seconds <= 15:
        raise ValueError("sample_seconds must be an integer in 1..15")
    target = Path(path)
    if not target.is_file():
        return failure("EXPORT_FILE_MISSING", "動画ファイルがありません（Python MCP側のローカルパスを指定）")
    try:
        if target.stat().st_size < 32:
            return failure("EXPORT_FILE_EMPTY", "動画ファイルが小さすぎます")
    except OSError:
        return failure("EXPORT_VERIFY_FAILED", "動画ファイルを読み取れません")

    try:
        probe = run(["ffprobe", "-v", "error", "-show_entries",
                     "format=duration:stream=index,codec_type,width,height,sample_rate,channels",
                     "-of", "json", str(target)], 15)
    except FileNotFoundError:
        return failure("FFMPEG_UNAVAILABLE", "ffprobeが見つかりません")
    except subprocess.TimeoutExpired:
        return failure("EXPORT_PROBE_TIMEOUT", "ffprobeが時間切れです")
    except OSError:
        return failure("EXPORT_PROBE_FAILED", "ffprobeを実行できません")
    if probe.returncode or len(probe.stdout) > 1_000_000:
        return failure("EXPORT_PROBE_FAILED", "ffprobeで動画を解析できません")
    try:
        info = json.loads(probe.stdout)
        streams = info["streams"]
        duration = float(info["format"]["duration"])
        if not isinstance(streams, list) or not math.isfinite(duration) or duration <= 0:
            raise ValueError("invalid metadata")
    except (ValueError, TypeError, KeyError, OverflowError):
        return failure("EXPORT_PROBE_FAILED", "動画の尺・ストリーム情報が不正です")

    videos = [s for s in streams if isinstance(s, dict) and s.get("codec_type") == "video"]
    audios = [s for s in streams if isinstance(s, dict) and s.get("codec_type") == "audio"]
    issues = []
    if not videos:
        issues.append({"code": "EXPORT_VIDEO_STREAM_MISSING", "severity": "error"})
    if require_audio and not audios:
        issues.append({"code": "EXPORT_AUDIO_STREAM_MISSING", "severity": "error"})
    starts = [0.0]
    if duration > 2 * sample_seconds:
        starts.append(max(0.0, duration - sample_seconds))
    samples = []
    black_frames = []
    long_silences = []
    audio_peaks = []
    for start in starts:
        length = min(float(sample_seconds), duration - start)
        if length <= 0:
            continue
        samples.append({"start_seconds": round(start, 3), "duration_seconds": round(length, 3)})
        for kind, present in (("video", bool(videos)), ("audio", bool(audios))):
            if not present:
                continue
            # Keep the decode check and add non-blocking signal hints. The
            # filters emit bounded info logs for the sampled windows only.
            signal_filter = ("blackdetect=d=1:pix_th=0.10:pic_th=0.98"
                             if kind == "video" else "silencedetect=noise=-50dB:d=2,astats=metadata=0:reset=0")
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-nostats",
                       "-v", "info", "-xerror", "-ss", format(start, ".3f"),
                       "-i", str(target), "-t", format(length, ".3f"),
                       "-map", "0:v:0" if kind == "video" else "0:a:0",
                       "-vf" if kind == "video" else "-af", signal_filter,
                       "-abort_on", "empty_output", "-f", "null", "-"]
            try:
                decoded = run(command, 30)
            except FileNotFoundError:
                return failure("FFMPEG_UNAVAILABLE", "ffmpegが見つかりません")
            except subprocess.TimeoutExpired:
                return failure("EXPORT_DECODE_TIMEOUT", "動画サンプルのデコードが時間切れです")
            except OSError:
                return failure("EXPORT_DECODE_FAILED", "ffmpegを実行できません")
            if decoded.returncode:
                issues.append({"code": "EXPORT_DECODE_FAILED", "severity": "error",
                               "stream_type": kind, "start_seconds": round(start, 3),
                               "details": (decoded.stderr or "")[-500:]})
                continue
            if kind == "audio":
                peak = _audio_peak_db(decoded.stderr or "")
                if peak is not None:
                    peak_event = {"sample_start_seconds": round(start, 3),
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
                    decoded.stderr or "", start_name, end_name):
                event = {"sample_start_seconds": round(start, 3),
                         "filter_start_seconds": round(signal_start, 3),
                         "filter_end_seconds": round(signal_end, 3),
                         "duration_seconds": round(signal_end - signal_start, 3)}
                if kind == "video":
                    black_frames.append(event)
                    code = "EXPORT_SAMPLE_BLACK_FRAMES"
                else:
                    long_silences.append(event)
                    code = "EXPORT_SAMPLE_LONG_SILENCE"
                issues.append({"code": code, "severity": "warning", **event})
    return {"success": True,
            "passed": not any(issue["severity"] == "error" for issue in issues),
            "issues": issues,
            "verification_scope": "sampled_audio_video_decode", "fully_decoded": False,
            "signal_scope": "sampled_only", "black_frames": black_frames,
            "long_silences": long_silences, "audio_peaks": audio_peaks,
            "duration_seconds": duration, "video_streams": len(videos),
            "audio_streams": len(audios), "samples": samples, "path": str(target)}
