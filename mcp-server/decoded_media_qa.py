"""Read-only, bounded FFmpeg decode QA for completed local video exports."""
import json
import math
import subprocess
from pathlib import Path

from jobs import is_absolute_media_path

VIDEO_EXTENSIONS = {".mp4", ".mov", ".mkv", ".webm", ".avi"}


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
    for start in starts:
        length = min(float(sample_seconds), duration - start)
        if length <= 0:
            continue
        samples.append({"start_seconds": round(start, 3), "duration_seconds": round(length, 3)})
        for kind, present in (("video", bool(videos)), ("audio", bool(audios))):
            if not present:
                continue
            command = ["ffmpeg", "-hide_banner", "-nostdin", "-v", "error", "-xerror",
                       "-ss", format(start, ".3f"), "-i", str(target),
                       "-t", format(length, ".3f"), "-map",
                       "0:v:0" if kind == "video" else "0:a:0", "-f", "null", "-"]
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
    return {"success": True, "passed": not issues, "issues": issues,
            "verification_scope": "sampled_audio_video_decode", "fully_decoded": False,
            "duration_seconds": duration, "video_streams": len(videos),
            "audio_streams": len(audios), "samples": samples, "path": str(target)}
