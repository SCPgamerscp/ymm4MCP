#!/usr/bin/env python3
"""
Gemini を使った動画解析モジュール。

YMM4 MCP サーバーから利用し、以下2つの解析対象に対応する:
  - 案A: YMM4プレビューを書き出した「連番PNG + 音声WAV」フォルダ
  - 案B: 取り込み前の元動画ファイル(mp4等)

検出するイベントの種類は AI(Claude) がプロンプトで自由に指定できる。
未指定時は「映像・音声で起きていることを全て」検出する。

戻り値は常に、タイムスタンプ(秒/ms)付きの構造化イベントリスト。
呼び出し側(server.py)が FPS を使ってフレーム番号へ変換する。

APIキー:
  環境変数 GEMINI_API_KEY または GOOGLE_API_KEY から取得する。
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
from typing import Any, Optional

import httpx
import requests

# google-genai は遅延 import する(未インストールでもサーバー自体は起動できるように)
_GENAI_IMPORT_ERROR: Optional[str] = None
try:
    from google import genai
    from google.genai import types as genai_types
except Exception as e:  # pragma: no cover
    genai = None
    genai_types = None
    _GENAI_IMPORT_ERROR = str(e)


DEFAULT_GEMINI_MODEL = "gemini-3.8-flash"
MAX_MODEL_CONFIG_BYTES = 64 * 1024
_MODEL_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9._/-]{0,127}")


def _unique_model_config_object(pairs: list[tuple[str, Any]]) -> dict:
    """Reject ambiguous model configuration keys, including nested metadata."""
    config = {}
    for key, value in pairs:
        if key in config:
            raise ValueError("duplicate Gemini model configuration key")
        config[key] = value
    return config


def _valid_model_id(value: Any) -> bool:
    """Allow explicit model names while rejecting ambiguous path segments."""
    return (isinstance(value, str) and _MODEL_ID.fullmatch(value) is not None
            and all(segment not in ("", ".", "..") for segment in value.split("/")))


def _get_default_model() -> str:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    try:
        if os.path.isfile(config_path):
            # A local corrupt/oversized config must not prevent video analysis.
            with open(config_path, "rb") as f:
                raw = f.read(MAX_MODEL_CONFIG_BYTES + 1)
            if len(raw) > MAX_MODEL_CONFIG_BYTES:
                return DEFAULT_GEMINI_MODEL
            config = json.loads(
                raw.decode("utf-8-sig"),
                object_pairs_hook=_unique_model_config_object,
            )
            if isinstance(config, dict):
                model = config.get("default_model")
                if _valid_model_id(model):
                    return model
    except (OSError, UnicodeError, ValueError):
        pass
    return DEFAULT_GEMINI_MODEL

DEFAULT_MODEL = _get_default_model()

# Limit memory retained for inline Gemini frame/audio payloads.
MAX_FRAME_IMAGE_BYTES = 32 * 1024 * 1024
MAX_AUDIO_BYTES = 128 * 1024 * 1024
MAX_TOTAL_BYTES = 256 * 1024 * 1024
# Gemini File API has a per-file size limit; avoid unbounded video uploads.
MAX_VIDEO_UPLOAD_BYTES = 2 * 1024 * 1024 * 1024

# デフォルトの検出指示(イベント全種類)。AIがカスタムプロンプトを渡せば上書きされる。
DEFAULT_EVENT_INSTRUCTION = (
    "この動画(または連番フレーム画像と音声)を時系列で精密に解析し、"
    "起きている出来事を可能な限りすべて検出してください。具体的には次を含みます:\n"
    "- シーンの切り替わり/カットの変化\n"
    "- 登場人物・キャラクター・オブジェクトの出現/退場\n"
    "- 大きな動きやアクション(攻撃、移動、衝突、変形 など)\n"
    "- 効果音・歓声・BGMの変化・無音→音ありの変化などの音響イベント\n"
    "- テロップ・文字・UI・数値(スコア/HP等)の表示や変化\n"
    "- 画面全体の色やフラッシュ、明暗の急変\n"
    "- その他、実況・解説の見どころになりそうな重要な瞬間"
)


def _get_api_key() -> Optional[str]:
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")


def is_available() -> tuple[bool, str]:
    """Gemini解析が利用可能か(SDK導入済み & APIキー設定済み)を返す。"""
    if genai is None:
        return False, f"google-genai 未インストール: {_GENAI_IMPORT_ERROR}. `pip install google-genai` を実行してください。"
    if not _get_api_key():
        return False, "環境変数 GEMINI_API_KEY (または GOOGLE_API_KEY) が未設定です。"
    return True, "ok"


def _client() -> "genai.Client":
    return genai.Client(api_key=_get_api_key())


def _build_prompt(event_instruction: str) -> str:
    """イベント検出指示を組み込んだ、JSON出力を強制するプロンプトを作る。"""
    return (
        f"{event_instruction}\n\n"
        "## 出力形式(厳守)\n"
        "解析結果を以下のJSON **のみ** で返してください。前後に説明文やコードフェンスを付けないこと。\n"
        "{\n"
        '  "summary": "動画全体の要約(日本語, 2〜4文)",\n'
        '  "events": [\n'
        "    {\n"
        '      "time_sec": 12.3,                // イベント発生時刻(秒, 小数可)。動画先頭=0。\n'
        '      "end_sec": 14.0,                 // 終了時刻(継続するイベントのみ。瞬間的ならtime_secと同じでよい)\n'
        '      "type": "scene_change | character | action | sound | text | color | other",\n'
        '      "label": "短いラベル(例: 敵キャラ出現)",\n'
        '      "description": "何が起きたかの説明(日本語)",\n'
        '      "audio": "その瞬間の音の説明(効果音/歓声/BGM/無音 など。不明ならnull)",\n'
        '      "importance": 1                  // 1(低)〜5(高) 実況の見どころ度\n'
        "    }\n"
        "  ]\n"
        "}\n\n"
        "重要: time_sec は必ず動画の先頭を0とした相対秒で答えること。events は時刻昇順で並べること。"
    )


MAX_GEMINI_RESPONSE_CHARS = 4 * 1024 * 1024
MAX_GEMINI_JSON_INTEGER_DIGITS = 128
MAX_GEMINI_JSON_DEPTH = 128


def _bounded_gemini_integer(token: str) -> int:
    """Bound integer parsing even on Python versions with a higher int limit."""
    if len(token.lstrip("-")) > MAX_GEMINI_JSON_INTEGER_DIGITS:
        raise ValueError("JSON integer literal is too long")
    return int(token)


def _gemini_json_too_deep(body: str) -> bool:
    """Count actual JSON containers, ignoring braces inside escaped strings."""
    depth = 0
    quoted = False
    escaped = False
    for char in body:
        if quoted:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                quoted = False
        elif char == '"':
            quoted = True
        elif char in "[{":
            depth += 1
            if depth > MAX_GEMINI_JSON_DEPTH:
                return True
        elif char in "]}":
            depth -= 1
    return False




def _unique_gemini_object(pairs: list[tuple[str, Any]]) -> dict:
    """Reject conflicting claims in nested model-generated JSON objects."""
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def _reject_gemini_constant(token: str) -> None:
    raise ValueError(f"nonstandard JSON numeric constant: {token}")


def _finite_gemini_float(token: str) -> float:
    if len(token) > 128:
        raise ValueError("JSON floating-point literal is too long")
    value = float(token)
    if not math.isfinite(value):
        raise ValueError("JSON floating-point number is not finite")
    return value


def _reject_gemini_surrogates(value: Any) -> None:
    """JSON escaped unpaired surrogates cannot be encoded as valid UTF-8."""
    pending = [value]
    while pending:
        current = pending.pop()
        if isinstance(current, str):
            if any(0xD800 <= ord(char) <= 0xDFFF for char in current):
                raise ValueError("unpaired Unicode surrogate in Gemini JSON")
        elif isinstance(current, dict):
            pending.extend(current.keys())
            pending.extend(current.values())
        elif isinstance(current, list):
            pending.extend(current)


def _extract_json(text: str) -> dict:
    """Geminiの応答からJSONを頑健に抽出する。コードフェンスや前後テキストを許容。"""
    if not text:
        return {"summary": "", "events": [], "_parse_error": "空の応答"}
    # The response has already been downloaded, but do not parse an unbounded
    # model response into a potentially huge object graph.
    if len(text) > MAX_GEMINI_RESPONSE_CHARS:
        return {"summary": "", "events": [], "_parse_error": "Gemini JSON response is too large"}
    # ```json ... ``` を剥がす
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = fence.group(1) if fence else None
    if candidate is None:
        # 最初の { から最後の } までを取る
        start = text.find("{")
        end = text.rfind("}")
        candidate = text[start : end + 1] if (start >= 0 and end > start) else text
    try:
        if _gemini_json_too_deep(candidate):
            raise ValueError("Gemini JSON exceeds the nesting depth limit")
        parsed = json.loads(
            candidate,
            object_pairs_hook=_unique_gemini_object,
            parse_constant=_reject_gemini_constant,
            parse_int=_bounded_gemini_integer,
            parse_float=_finite_gemini_float,
        )
        _reject_gemini_surrogates(parsed)
        if (not isinstance(parsed, dict)
                or not isinstance(parsed.get("summary"), str)
                or not isinstance(parsed.get("events"), list)
                or any(not isinstance(event, dict) for event in parsed["events"])):
            raise ValueError("summary must be text and events must be an array of objects")
        if len(parsed["events"]) > 5000:
            raise ValueError("events exceeds the 5000 event limit")
        for index, event in enumerate(parsed["events"]):
            start = event.get("time_sec")
            if (isinstance(start, bool) or not isinstance(start, (int, float))
                    or not math.isfinite(start) or start < 0):
                raise ValueError(f"events[{index}].time_sec must be finite and nonnegative")
            end = event.get("end_sec", start)
            if end is not None and (
                    isinstance(end, bool) or not isinstance(end, (int, float))
                    or not math.isfinite(end) or end < start):
                raise ValueError(f"events[{index}].end_sec must not precede time_sec")
        # Gemini can return valid events out of order. Preserve relative order
        # for ties while guaranteeing the chronological contract for callers.
        parsed["events"].sort(key=lambda event: event["time_sec"])
        return parsed
    except Exception as e:
        return {"summary": "", "events": [], "_parse_error": f"JSON解析失敗: {e}", "_raw": text[:2000]}


def _video_signature(path: str) -> tuple[int, int, int, int]:
    """Detect replacement or mutation of the file being analyzed."""
    info = os.stat(path)
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns


def _media_still_matches(snapshots: list[tuple[str, tuple[int, int, int, int], bytes]]) -> bool:
    """Recheck selected preview media after sampling and after model inference.

    Stat alone cannot detect an in-place rewrite that restores the timestamp.
    Compare both file identity and a streaming digest of the exact bytes sent.
    """
    for path, original, digest in snapshots:
        try:
            if _video_signature(path) != original:
                return False
            hasher = hashlib.sha256()
            with open(path, "rb") as source:
                for chunk in iter(lambda: source.read(1024 * 1024), b""):
                    hasher.update(chunk)
            if _video_signature(path) != original or hasher.digest() != digest:
                return False
        except OSError:
            return False
    return True


def _transient_file_api_error(exc: Exception) -> bool:
    """Only retry temporary File API failures, never auth or bad requests."""
    return (getattr(exc, "code", None) in (429, 500, 502, 503, 504)
            or getattr(exc, "status_code", None) in (429, 500, 502, 503, 504)
            or isinstance(exc, (TimeoutError, ConnectionError, httpx.TransportError,
                                requests.exceptions.Timeout,
                                requests.exceptions.ConnectionError)))


def _wait_file_active(client, file_obj, timeout_sec: int = 180):
    """File API にアップロードした動画が ACTIVE になるまで待つ。"""
    name = file_obj.name
    deadline = time.monotonic() + timeout_sec
    consecutive_errors = 0
    while True:
        # SDK v0.3.0 silently ignored HTTP timeouts. Since v1.0.0,
        # GetFileConfig.http_options.timeout is passed to the HTTP transport
        # in milliseconds. Bound each request by the *remaining* budget,
        # rather than checking the deadline only after a blocked call returns.
        remaining_ms = int((deadline - time.monotonic()) * 1000)
        if remaining_ms < 1:
            raise TimeoutError("Gemini File API polling timed out")
        try:
            f = client.files.get(
                name=name, config={"http_options": {"timeout": remaining_ms}}
            )
        except Exception as exc:
            if not _transient_file_api_error(exc):
                raise
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError("Gemini File API polling timed out") from exc
            # Back off under 429/5xx, but never sleep beyond the same deadline.
            consecutive_errors += 1
            time.sleep(min(3 * (2 ** min(consecutive_errors - 1, 3)), remaining))
            continue
        # Even an ACTIVE response is too late if the network call exceeded
        # the polling budget; otherwise it can incorrectly certify stale QA.
        if time.monotonic() >= deadline:
            raise TimeoutError("Gemini File API polling timed out")
        consecutive_errors = 0
        state = getattr(f.state, "name", str(f.state))
        if state == "ACTIVE":
            return f
        if state == "FAILED":
            raise RuntimeError("Geminiでの動画処理に失敗しました(FAILED)")
        # Account for the time spent inside files.get, not just sleep calls.
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise TimeoutError(f"動画処理がタイムアウトしました(state={state})")
        time.sleep(min(3, remaining))


def analyze_video_file(
    video_path: str,
    event_instruction: Optional[str] = None,
    model: str = DEFAULT_MODEL,
) -> dict:
    """
    【案B】元動画ファイル(mp4等)を直接 Gemini で解析する。

    Gemini の File API に動画をアップロードし、映像+音声を時系列解析させて
    タイムスタンプ付きイベントJSONを得る。
    """
    if (not isinstance(video_path, str) or not video_path.strip()
            or any(ord(char) < 32 or ord(char) == 127 for char in video_path)):
        return {"success": False, "error": "video_path must be a nonempty path without control characters"}
    ok, msg = is_available()
    if not ok:
        return {"success": False, "error": msg}
    if not os.path.isfile(video_path):
        return {"success": False, "error": f"動画ファイルが見つかりません: {video_path}"}
    try:
        signature = _video_signature(video_path)
    except OSError as exc:
        return {"success": False, "error": f"動画ファイルを確認できません: {exc}"}
    if signature[2] <= 0 or signature[2] > MAX_VIDEO_UPLOAD_BYTES:
        return {"success": False, "error": "Video file is empty or exceeds the Gemini upload size limit"}

    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION
    client = None
    uploaded = None
    try:
        client = _client()
        uploaded = client.files.upload(file=video_path)
        # A file modified during upload may not match the local media that
        # the QA report will be attributed to. Never certify such a result.
        if _video_signature(video_path) != signature:
            return {"success": False, "error": "Video changed during Gemini upload", "source": "file"}
        active = _wait_file_active(client, uploaded)
        resp = client.models.generate_content(
            model=model,
            contents=[active, _build_prompt(instruction)],
        )
        if _video_signature(video_path) != signature:
            return {"success": False, "error": "Video changed during Gemini analysis", "source": "file"}
        parsed = _extract_json(getattr(resp, "text", "") or "")
        if "_parse_error" in parsed:
            return {"success": False, "error": parsed["_parse_error"], "source": "file"}
        return {
            **parsed,
            "success": True,
            "source": "file",
            "video_path": video_path,
            "model": model,
        }
    except Exception as e:
        return {"success": False, "error": f"Gemini解析エラー: {e}"}
    finally:
        # File API uploads must be removed even when processing or parsing fails.
        if client is not None and uploaded is not None:
            try:
                client.files.delete(name=uploaded.name)
            except Exception:
                pass


def _safe_frame_path(frames_dir: str, filename: Any) -> Optional[str]:
    """Allow only existing PNG frames contained in the export directory.

    Resolve symlinks before checking containment, so untrusted export metadata
    cannot make the Gemini client read and upload files outside the export.
    """
    if not isinstance(filename, str) or not filename.lower().endswith(".png"):
        return None
    root = os.path.realpath(frames_dir)
    path = os.path.realpath(os.path.join(root, filename))
    try:
        if os.path.commonpath((root, path)) != root:
            return None
    except ValueError:  # Different drives on Windows.
        return None
    return path if os.path.isfile(path) else None


def _frame_filename_key(name: Any) -> tuple:
    """Natural numeric ordering for exports such as 1.png, 2.png, 10.png."""
    if not isinstance(name, str):
        return ((0, ""),)
    return tuple((1, int(part)) if part.isdigit() else (0, part.casefold())
                 for part in re.split(r"([0-9]+)", name))


def _valid_time_ms(value: Any) -> bool:
    return (type(value) in (int, float) and 0 <= value <= 1_000_000_000_000
            and math.isfinite(value))


def _valid_frame_index(value: Any) -> bool:
    return type(value) is int and 0 <= value <= 2_147_483_647


def analyze_frames_dir(
    frames_dir: str,
    frames_meta: Optional[list[dict]] = None,
    audio_path: Optional[str] = None,
    event_instruction: Optional[str] = None,
    model: str = DEFAULT_MODEL,
    max_images: int = 60,
) -> dict:
    """
    【案A】YMM4プレビューを書き出した連番PNG(+音声WAV)フォルダを Gemini で解析する。

    frames_meta は C# の /api/preview/export-clip が返す frames 配列
    (各要素 {index, frame, timeMs, file, ...})。time_sec とフレーム番号の対応に使う。
    画像が多すぎる場合は均等間引きして max_images 枚に絞る。
    """
    ok, msg = is_available()
    if not ok:
        return {"success": False, "error": msg}
    if not os.path.isdir(frames_dir):
        return {"success": False, "error": f"フォルダが見つかりません: {frames_dir}"}

    if (isinstance(max_images, bool) or not isinstance(max_images, int)
            or not 1 <= max_images <= 1000):
        return {"success": False, "error": "max_images must be an integer in 1..1000"}
    if frames_meta is not None and (not isinstance(frames_meta, list)
                                    or any(not isinstance(m, dict) for m in frames_meta)):
        return {"success": False, "error": "frames_meta must be an array of objects"}

    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION

    # 解析対象の画像ファイルを決定(metaがあればそれを優先、なければフォルダ走査)
    if frames_meta:
        entries = [m for m in frames_meta if m.get("file")]
        if len(entries) > 1:
            # Never claim chronological analysis for shuffled or incomparable
            # host metadata. Prefer actual capture times, then frame indices.
            if all(_valid_time_ms(e.get("timeMs")) for e in entries):
                entries.sort(key=lambda e: e["timeMs"])
            elif all(_valid_frame_index(e.get("frame")) for e in entries):
                entries.sort(key=lambda e: e["frame"])
            elif any(_valid_time_ms(e.get("timeMs")) or _valid_frame_index(e.get("frame"))
                     for e in entries):
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta has inconsistent timestamp/frame metadata"}
            else:
                entries.sort(key=lambda e: _frame_filename_key(e["file"]))
    else:
        files = sorted((f for f in os.listdir(frames_dir) if f.lower().endswith(".png")),
                       key=_frame_filename_key)
        entries = [{"file": f, "timeMs": None, "frame": None} for f in files]

    if not entries:
        return {"success": False, "error": "解析対象の画像がありません"}

    # 多すぎる場合は均等間引き
    if len(entries) > max_images:
        stepf = len(entries) / max_images
        entries = [entries[int(i * stepf)] for i in range(max_images)]

    # 画像 + 各画像の時刻ラベルを contents に並べる
    contents: list[Any] = []
    time_table_lines = []
    try:
        client = _client()
        total_payload_bytes = 0
        analyzed_images = 0
        media_snapshots: list[tuple[str, tuple[int, int, int, int], bytes]] = []
        for e in entries:
            path = _safe_frame_path(frames_dir, e["file"])
            if path is None:
                # A partial frame sequence can make Gemini miss an actual
                # problem while still returning success; never silently skip.
                return {"success": False, "source": "frames_dir",
                        "error": "Selected PNG frame is missing or outside export directory"}
            if os.path.getsize(path) > MAX_FRAME_IMAGE_BYTES:
                return {"success": False, "error": "PNG frame exceeds the upload size limit",
                        "source": "frames_dir"}
            frame_signature = _video_signature(path)
            with open(path, "rb") as fh:
                img_bytes = fh.read(MAX_FRAME_IMAGE_BYTES + 1)
            if (len(img_bytes) > MAX_FRAME_IMAGE_BYTES
                    or total_payload_bytes + len(img_bytes) > MAX_TOTAL_BYTES):
                return {"success": False, "error": "Frame payload exceeds the upload size limit",
                        "source": "frames_dir"}
            if (_video_signature(path) != frame_signature
                    or len(img_bytes) != frame_signature[2]):
                return {"success": False, "source": "frames_dir",
                        "error": "Selected PNG frame changed while being read"}
            media_snapshots.append(
                (path, frame_signature, hashlib.sha256(img_bytes).digest()))
            total_payload_bytes += len(img_bytes)
            contents.append(genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png"))
            analyzed_images += 1
            t_ms = e.get("timeMs")
            valid_time = (not isinstance(t_ms, bool) and isinstance(t_ms, (int, float))
                          and 0 <= t_ms <= 1_000_000_000_000 and math.isfinite(t_ms))
            t_label = f"{t_ms/1000.0:.2f}s" if valid_time else "?"
            frame = e.get("frame")
            frame_label = (frame if type(frame) is int and 0 <= frame <= 2_147_483_647
                           else "?")
            contents.append(genai_types.Part.from_text(text=f"[上の画像の時刻 = {t_label}, frame={frame_label}]"))
            time_table_lines.append(f"  画像{len(time_table_lines)}: time={t_label} frame={frame_label}")

        if not time_table_lines:
            return {"success": False, "error": "解析可能なPNGフレームがありません",
                    "source": "frames_dir"}

        # An explicitly requested audio track is required for audiovisual QA.
        # Silently omitting a missing WAV would falsely report a complete review.
        if audio_path is not None:
            if (not isinstance(audio_path, str) or not audio_path or
                    not audio_path.lower().endswith(".wav") or not os.path.isfile(audio_path)):
                return {"success": False, "source": "frames_dir",
                        "error": "Requested audio file must be an existing WAV"}
            if os.path.getsize(audio_path) > MAX_AUDIO_BYTES:
                return {"success": False, "error": "Audio exceeds the upload size limit",
                        "source": "frames_dir"}
            audio_signature = _video_signature(audio_path)
            with open(audio_path, "rb") as fh:
                wav_bytes = fh.read(MAX_AUDIO_BYTES + 1)
            if (len(wav_bytes) > MAX_AUDIO_BYTES
                    or total_payload_bytes + len(wav_bytes) > MAX_TOTAL_BYTES):
                return {"success": False, "error": "Audio payload exceeds the upload size limit",
                        "source": "frames_dir"}
            # Do not label arbitrary audio bytes as audio/wav in Gemini requests.
            if (len(wav_bytes) < 12 or wav_bytes[:4] != b"RIFF"
                    or wav_bytes[8:12] != b"WAVE"):
                return {"success": False, "error": "Requested audio is not a RIFF/WAVE file",
                        "source": "frames_dir"}
            if (_video_signature(audio_path) != audio_signature
                    or len(wav_bytes) != audio_signature[2]):
                return {"success": False, "source": "frames_dir",
                        "error": "Requested audio changed while being read"}
            media_snapshots.append(
                (audio_path, audio_signature, hashlib.sha256(wav_bytes).digest()))
            total_payload_bytes += len(wav_bytes)
            contents.append(genai_types.Part.from_bytes(data=wav_bytes, mime_type="audio/wav"))
            contents.append(genai_types.Part.from_text(
                text="↑これは上記フレーム区間全体の音声です。画像の時刻と対応付けて音響イベントを解析してください。"))

        # 指示プロンプト
        contents.append(genai_types.Part.from_text(text=(
            "上記は、ある動画区間を時刻順にサンプリングした連番フレーム画像"
            "(各画像の直後にその時刻を示すラベルあり)と、その区間の音声です。\n\n"
            + _build_prompt(instruction)
        )))

        if not _media_still_matches(media_snapshots):
            return {"success": False, "source": "frames_dir",
                    "error": "Preview media changed before Gemini analysis"}
        resp = client.models.generate_content(model=model, contents=contents)
        if not _media_still_matches(media_snapshots):
            return {"success": False, "source": "frames_dir",
                    "error": "Preview media changed during Gemini analysis"}
        parsed = _extract_json(getattr(resp, "text", "") or "")
        if "_parse_error" in parsed:
            return {"success": False, "error": parsed["_parse_error"], "source": "frames_dir"}
        return {
            **parsed,
            "success": True,
            "source": "frames_dir",
            "frames_dir": frames_dir,
            "analyzed_images": analyzed_images,
            "model": model,
        }
    except Exception as e:
        return {"success": False, "error": f"Gemini解析エラー: {e}"}
