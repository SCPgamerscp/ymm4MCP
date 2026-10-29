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

import json
import math
import os
import re
import time
from typing import Any, Optional

# google-genai は遅延 import する(未インストールでもサーバー自体は起動できるように)
_GENAI_IMPORT_ERROR: Optional[str] = None
try:
    from google import genai
    from google.genai import types as genai_types
except Exception as e:  # pragma: no cover
    genai = None
    genai_types = None
    _GENAI_IMPORT_ERROR = str(e)


def _get_default_model() -> str:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    try:
        if os.path.isfile(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f).get("default_model", "gemini-2.5-flash")
    except Exception:
        pass
    return "gemini-2.5-flash"

DEFAULT_MODEL = _get_default_model()

# Limit memory retained for inline Gemini frame/audio payloads.
MAX_FRAME_IMAGE_BYTES = 32 * 1024 * 1024
MAX_AUDIO_BYTES = 128 * 1024 * 1024
MAX_TOTAL_BYTES = 256 * 1024 * 1024

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


def _extract_json(text: str) -> dict:
    """Geminiの応答からJSONを頑健に抽出する。コードフェンスや前後テキストを許容。"""
    if not isinstance(text, str):
        return {"summary": "", "events": [], "_parse_error": "Gemini response must be text"}
    # A malformed or unexpectedly verbose model response must not allocate
    # unbounded JSON objects while automated video QA is running.
    if len(text) > 16 * 1024 * 1024:
        return {"summary": "", "events": [], "_parse_error": "Gemini response exceeds 16777216 characters"}
    if not text:
        return {"summary": "", "events": [], "_parse_error": "空の応答"}
    # ```json ... ``` を剥がす
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = fence.group(1) if fence else None
    if candidate is None:
        # 最初の { から最後の } までを取る
        start = text.find("{")
        end = text.rfind("}")
        candidate = text[start : end + 1] if (start >= 0 and end > start) else text
    try:
        parsed = json.loads(candidate)
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


def _wait_file_active(client, file_obj, timeout_sec: int = 180):
    """File API にアップロードした動画が ACTIVE になるまで待つ。"""
    name = file_obj.name
    deadline = time.monotonic() + timeout_sec
    while True:
        f = client.files.get(name=name)
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
    if event_instruction is not None and not _valid_event_instruction(event_instruction):
        return {"success": False, "source": "file",
                "error": "event_instruction must be valid text of at most 32768 characters"}
    ok, msg = is_available()
    if not ok:
        return {"success": False, "error": msg}
    if not os.path.isfile(video_path):
        return {"success": False, "error": f"動画ファイルが見つかりません: {video_path}"}

    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION
    client = None
    uploaded = None
    try:
        client = _client()
        uploaded = client.files.upload(file=video_path)
        active = _wait_file_active(client, uploaded)
        resp = client.models.generate_content(
            model=model,
            contents=[active, _build_prompt(instruction)],
        )
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


def _valid_event_instruction(value: Any) -> bool:
    """Reject unbounded or non-UTF-8-encodable user-supplied Gemini prompts."""
    return (isinstance(value, str) and len(value) <= 32_768
            and not any(ord(ch) == 0 or 0xD800 <= ord(ch) <= 0xDFFF
                        for ch in value))


def _valid_capture_png_name(name: Any) -> bool:
    """Accept only unambiguous Windows-exportable PNG basenames.

    YMM4 exports on Windows. Reserved DOS device names and alternate data
    stream separators are unsafe even if a POSIX test host can create them.
    """
    if (not isinstance(name, str) or not name or len(name) > 255
            or name.strip() != name or not name.lower().endswith(".png")
            or any(ord(ch) < 32 or ord(ch) == 127
                   or ch in '<>:"/\\|?*'
                   or 0xD800 <= ord(ch) <= 0xDFFF
                   or ord(ch) in (0x061C, 0x200E, 0x200F)
                   or 0x202A <= ord(ch) <= 0x202E
                   or 0x2066 <= ord(ch) <= 0x206F for ch in name)):
        return False
    stem = name.split(".", 1)[0].upper()
    return stem not in {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} and not (
        len(stem) == 4 and stem[:3] in ("COM", "LPT")
        and stem[3] in "123456789¹²³"
    )


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
    # Avoid int()'s digit limit and non-decimal Unicode isdigit() cases.
    def key(part: str) -> tuple:
        if part.isascii() and part.isdigit():
            normalized = part.lstrip("0") or "0"
            return (1, len(normalized), normalized, len(part))
        return (0, part.casefold(), "", 0)

    return tuple(key(part) for part in re.split(r"([0-9]+)", name))


def _valid_time_ms(value: Any) -> bool:
    return (type(value) in (int, float) and 0 <= value <= 1_000_000_000_000
            and math.isfinite(value))


def _valid_frame_index(value: Any) -> bool:
    return type(value) is int and 0 <= value <= 2_147_483_647


def _file_snapshot(info) -> tuple:
    """Capture file identity and metadata to detect mid-QA replacements."""
    return (info.st_dev, info.st_ino, info.st_size,
            info.st_mtime_ns, info.st_ctime_ns)


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
    # Even though only max_images frames are uploaded, sorting a huge metadata
    # list before sampling could otherwise exhaust memory or CPU.
    if frames_meta is not None and len(frames_meta) > 100_000:
        return {"success": False, "source": "frames_dir",
                "error": "frames_meta exceeds the 100000 frame metadata limit"}
    # Explicit empty host capture metadata must not fall back to stale PNGs
    # left in the export directory by an earlier capture.
    if frames_meta is not None and not frames_meta:
        return {"success": False, "source": "frames_dir",
                "error": "frames_meta contains no captured frames"}

    # A caller-supplied QA instruction must not create an unbounded prompt
    # or carry malformed Unicode into the Gemini request.
    if event_instruction is not None and not _valid_event_instruction(event_instruction):
        return {"success": False, "source": "frames_dir",
                "error": "event_instruction must be valid text of at most 32768 characters"}
    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION

    # 解析対象の画像ファイルを決定(metaがあればそれを優先、なければフォルダ走査)
    if frames_meta:
        # A partial selection must never be reported as a complete QA pass.
        # The YMM4 export API emits plain basenames. Reject path aliases,
        # traversal, control characters and padded names before any disk I/O.
        if any(not _valid_capture_png_name(m.get("file")) for m in frames_meta):
            return {"success": False, "source": "frames_dir",
                    "error": "frames_meta contains an invalid PNG filename"}
        entries = list(frames_meta)
        if len({e["file"] for e in entries}) != len(entries):
            return {"success": False, "source": "frames_dir",
                    "error": "frames_meta contains duplicate PNG filenames"}
        # Captures are produced by YMM4 on Windows. Case-only aliases can
        # resolve to the same file there even if tests run on a POSIX host.
        if len({e["file"].casefold() for e in entries}) != len(entries):
            return {"success": False, "source": "frames_dir",
                    "error": "frames_meta contains case-insensitive duplicate PNG filenames"}
        # The YMM4 exporter numbers captures from zero. An incomplete or
        # duplicated index sequence can conceal missing frames even when all
        # remaining filenames and timestamps look valid.
        if any("index" in e for e in entries):
            if (any(type(e.get("index")) is not int or e["index"] < 0
                    for e in entries)
                    or {e["index"] for e in entries} != set(range(len(entries)))):
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta contains incomplete capture indices"}
        # Older callers may omit dimensions, but if the exporter supplies them
        # they must be complete and usable for every captured image.
        if any("width" in e or "height" in e for e in entries):
            if any(type(e.get(axis)) is not int or not 1 <= e[axis] <= 32_768
                   for e in entries for axis in ("width", "height")):
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta contains invalid capture dimensions"}
            # A single export-clip capture has one fixed preview size.
            # Mixing dimensions means the capture changed mid-export and
            # cannot be treated as a consistent video QA sample.
            if len({(e["width"], e["height"]) for e in entries}) != 1:
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta contains inconsistent capture dimensions"}
        time_present = ["timeMs" in e for e in entries]
        frame_present = ["frame" in e for e in entries]
        has_time = [_valid_time_ms(e.get("timeMs")) for e in entries]
        has_frame = [_valid_frame_index(e.get("frame")) for e in entries]
        # A present-but-invalid field is not the same as an absent field.
        # All-invalid metadata must fail instead of falling back to filenames.
        if ((any(time_present) and (not all(time_present) or not all(has_time)))
                or (any(frame_present) and (not all(frame_present) or not all(has_frame)))):
            return {"success": False, "source": "frames_dir",
                    "error": "frames_meta has inconsistent timestamp/frame metadata"}
        if all(has_time):
            entries.sort(key=lambda e: e["timeMs"])
        elif all(has_frame):
            entries.sort(key=lambda e: e["frame"])
        elif all("index" in e for e in entries):
            # Without frame/time metadata, capture indices are authoritative.
            # Filename sorting can reverse the actual capture order.
            entries.sort(key=lambda e: e["index"])
        else:
            entries.sort(key=lambda e: _frame_filename_key(e["file"]))
        # The exporter's capture index is chronological. Even individually
        # valid indices must agree with the timestamp/frame sequence.
        if (len(entries) > 1 and all("index" in e for e in entries)
                and (all(has_time) or all(has_frame))
                and [e["index"] for e in entries] != list(range(len(entries)))):
            return {"success": False, "source": "frames_dir",
                    "error": "frames_meta capture indices conflict with chronological order"}
        # Timestamp order and frame order must agree when both are supplied.
        if len(entries) > 1:
            if all(has_time) and any(entries[i]["timeMs"] <= entries[i - 1]["timeMs"]
                                     for i in range(1, len(entries))):
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta contains duplicate capture timestamps"}
            if all(has_frame) and any(entries[i]["frame"] <= entries[i - 1]["frame"]
                                      for i in range(1, len(entries))):
                return {"success": False, "source": "frames_dir",
                        "error": "frames_meta frame order conflicts with capture order"}
    else:
        try:
            # scandir is incremental; listdir materializes every directory
            # entry before a limit can be applied. Bound even non-PNG entries
            # so an untrusted directory cannot starve the QA process.
            files = []
            with os.scandir(frames_dir) as iterator:
                for scanned_count, item in enumerate(iterator, start=1):
                    if scanned_count > 100_000:
                        return {"success": False, "source": "frames_dir",
                                "error": "PNG directory exceeds the 100000 entry scan limit"}
                    if item.name.lower().endswith(".png"):
                        if not _valid_capture_png_name(item.name):
                            return {"success": False, "source": "frames_dir",
                                    "error": "PNG directory contains an invalid filename"}
                        files.append(item.name)
            # Case-only aliases can be distinct on a POSIX test host but
            # collide when YMM4's Windows export is opened again.
            if len({name.casefold() for name in files}) != len(files):
                return {"success": False, "source": "frames_dir",
                        "error": "PNG directory contains case-insensitive duplicate filenames"}
            files.sort(key=_frame_filename_key)
        except OSError as exc:
            return {"success": False, "source": "frames_dir",
                    "error": f"Cannot enumerate PNG frames: {exc}"}
        entries = [{"file": f, "timeMs": None, "frame": None} for f in files]

    if not entries:
        return {"success": False, "error": "解析対象の画像がありません"}

    # Check the complete capture before downsampling, even without explicit
    # metadata. A broken unselected PNG must not yield a successful QA result.
    seen_capture_paths = set()
    seen_capture_file_ids = set()
    preflight_snapshots = {}
    for e in entries:
        path = _safe_frame_path(frames_dir, e["file"])
        if path is None:
            return {"success": False, "source": "frames_dir",
                    "error": "PNG frame is missing or outside export directory"}
        # Distinct names (including symlink aliases) may refer to the same
        # underlying frame. Treat such metadata as an incomplete capture.
        canonical = os.path.normcase(path)
        if canonical in seen_capture_paths:
            return {"success": False, "source": "frames_dir",
                    "error": "capture contains duplicate resolved PNG frames"}
        seen_capture_paths.add(canonical)
        try:
            info = os.stat(path)
        except OSError as exc:
            return {"success": False, "source": "frames_dir",
                    "error": f"Cannot inspect PNG frame: {exc}"}
        # Different hardlink names have different resolved paths but the
        # same file identity. They must not count as independent captures.
        if info.st_ino:
            file_id = (info.st_dev, info.st_ino)
            if file_id in seen_capture_file_ids:
                return {"success": False, "source": "frames_dir",
                        "error": "capture contains duplicate physical PNG frames"}
            seen_capture_file_ids.add(file_id)
        # Preserve the identity and metadata of every frame checked during
        # preflight. Selected files must still be the same immutable capture
        # when their bytes are later read for the Gemini request.
        preflight_snapshots[canonical] = _file_snapshot(info)
        # Keep the existing size-check seam used by preflight regression tests.
        try:
            size = os.path.getsize(path)
        except OSError as exc:
            return {"success": False, "source": "frames_dir",
                    "error": f"Cannot inspect PNG frame: {exc}"}
        if size == 0:
            return {"success": False, "source": "frames_dir",
                    "error": "PNG frame is empty"}
        if size > MAX_FRAME_IMAGE_BYTES:
            return {"success": False, "source": "frames_dir",
                    "error": "PNG frame exceeds the upload size limit"}

    # 多すぎる場合は均等間引き
    if len(entries) > max_images:
        # Cover the final frame as well as the first. The previous floor(i*N/K)
        # sampling always skipped the last frame when downsampling, potentially
        # missing a broken export tail. Use integer arithmetic for stable indices.
        if max_images == 1:
            entries = [entries[0]]
        else:
            last = len(entries) - 1
            entries = [entries[(i * last) // (max_images - 1)]
                       for i in range(max_images)]

    # 画像 + 各画像の時刻ラベルを contents に並べる
    contents: list[Any] = []
    time_table_lines = []
    try:
        client = _client()
        total_payload_bytes = 0
        analyzed_images = 0
        for e in entries:
            path = _safe_frame_path(frames_dir, e["file"])
            if path is None:
                # A partial frame sequence can make Gemini miss an actual
                # problem while still returning success; never silently skip.
                return {"success": False, "source": "frames_dir",
                        "error": "Selected PNG frame is missing or outside export directory"}
            canonical = os.path.normcase(path)
            expected = preflight_snapshots.get(canonical)
            if expected is None:
                return {"success": False, "source": "frames_dir",
                        "error": "Selected PNG frame changed after preflight"}
            # O_NOFOLLOW prevents a last-component symlink swap where supported;
            # fstat and path stat verify the actual opened file and its contents.
            flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
            fd = os.open(path, flags)
            with os.fdopen(fd, "rb") as fh:
                before = _file_snapshot(os.fstat(fh.fileno()))
                if before != expected:
                    return {"success": False, "source": "frames_dir",
                            "error": "Selected PNG frame changed after preflight"}
                img_bytes = fh.read(MAX_FRAME_IMAGE_BYTES + 1)
                after = _file_snapshot(os.fstat(fh.fileno()))
            if (after != expected or _file_snapshot(os.stat(path)) != expected):
                return {"success": False, "source": "frames_dir",
                        "error": "Selected PNG frame changed during upload preparation"}
            if not img_bytes:
                return {"success": False, "error": "Selected PNG frame is empty",
                        "source": "frames_dir"}
            if (len(img_bytes) > MAX_FRAME_IMAGE_BYTES
                    or total_payload_bytes + len(img_bytes) > MAX_TOTAL_BYTES):
                return {"success": False, "error": "Frame payload exceeds the upload size limit",
                        "source": "frames_dir"}
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
            # A RIFF/WAVE signature alone is insufficient: a truncated WAV
            # can still start with RIFF and WAVE. Do not certify an incomplete
            # audio track or upload appended bytes outside the RIFF container.
            riff_size = int.from_bytes(wav_bytes[4:8], "little")
            if riff_size < 4 or riff_size != len(wav_bytes) - 8:
                return {"success": False, "source": "frames_dir",
                        "error": "Requested audio RIFF/WAVE size does not match file length"}
            total_payload_bytes += len(wav_bytes)
            contents.append(genai_types.Part.from_bytes(data=wav_bytes, mime_type="audio/wav"))
            contents.append(genai_types.Part.from_text(
                text="↑これは上記フレーム区間全体の音声です。画像の時刻と対応付けて音響イベントを解析してください。"))

        # 指示プロンプト
        media_description = (
            "上記は、ある動画区間を時刻順にサンプリングした連番フレーム画像"
            "(各画像の直後にその時刻を示すラベルあり)"
        )
        if audio_path is not None:
            media_description += "と、その区間の音声です。"
        else:
            media_description += "です。音声データは提供されていません。"
        prompt = media_description + "\n\n" + _build_prompt(instruction)
        if audio_path is None:
            prompt += (
                "\n\n重要: 音声データがないため、BGM・効果音・声・無音などの"
                "音響イベントを観測したと推測して報告しないでください。"
            )
        contents.append(genai_types.Part.from_text(text=prompt))

        resp = client.models.generate_content(model=model, contents=contents)
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
