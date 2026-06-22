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


DEFAULT_MODEL = "gemini-2.5-flash"

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
        return json.loads(candidate)
    except Exception as e:
        return {"summary": "", "events": [], "_parse_error": f"JSON解析失敗: {e}", "_raw": text[:2000]}


def _wait_file_active(client, file_obj, timeout_sec: int = 180):
    """File API にアップロードした動画が ACTIVE になるまで待つ。"""
    name = file_obj.name
    waited = 0
    while True:
        f = client.files.get(name=name)
        state = getattr(f.state, "name", str(f.state))
        if state == "ACTIVE":
            return f
        if state == "FAILED":
            raise RuntimeError("Geminiでの動画処理に失敗しました(FAILED)")
        if waited >= timeout_sec:
            raise TimeoutError(f"動画処理がタイムアウトしました(state={state})")
        time.sleep(3)
        waited += 3


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
    ok, msg = is_available()
    if not ok:
        return {"success": False, "error": msg}
    if not os.path.isfile(video_path):
        return {"success": False, "error": f"動画ファイルが見つかりません: {video_path}"}

    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION
    client = _client()

    try:
        uploaded = client.files.upload(file=video_path)
        active = _wait_file_active(client, uploaded)
        resp = client.models.generate_content(
            model=model,
            contents=[active, _build_prompt(instruction)],
        )
        parsed = _extract_json(getattr(resp, "text", "") or "")
        # 後始末(File APIの容量節約)
        try:
            client.files.delete(name=uploaded.name)
        except Exception:
            pass
        return {
            "success": True,
            "source": "file",
            "video_path": video_path,
            "model": model,
            **parsed,
        }
    except Exception as e:
        return {"success": False, "error": f"Gemini解析エラー: {e}"}


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

    instruction = event_instruction or DEFAULT_EVENT_INSTRUCTION
    client = _client()

    # 解析対象の画像ファイルを決定(metaがあればそれを優先、なければフォルダ走査)
    if frames_meta:
        entries = [m for m in frames_meta if m.get("file")]
    else:
        files = sorted(f for f in os.listdir(frames_dir) if f.lower().endswith(".png"))
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
        for e in entries:
            path = os.path.join(frames_dir, e["file"])
            if not os.path.isfile(path):
                continue
            with open(path, "rb") as fh:
                img_bytes = fh.read()
            contents.append(genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png"))
            t_ms = e.get("timeMs")
            t_label = f"{t_ms/1000.0:.2f}s" if isinstance(t_ms, (int, float)) else "?"
            contents.append(genai_types.Part.from_text(text=f"[上の画像の時刻 = {t_label}, frame={e.get('frame')}]"))
            time_table_lines.append(f"  画像{len(time_table_lines)}: time={t_label} frame={e.get('frame')}")

        # 音声があれば添付
        if audio_path and os.path.isfile(audio_path):
            with open(audio_path, "rb") as fh:
                wav_bytes = fh.read()
            contents.append(genai_types.Part.from_bytes(data=wav_bytes, mime_type="audio/wav"))
            contents.append(genai_types.Part.from_text(
                text="↑これは上記フレーム区間全体の音声です。画像の時刻と対応付けて音響イベントを解析してください。"))

        # 指示プロンプト
        contents.append(genai_types.Part.from_text(text=(
            "上記は、ある動画区間を時刻順にサンプリングした連番フレーム画像"
            "(各画像の直後にその時刻を示すラベルあり)と、その区間の音声です。\n\n"
            + _build_prompt(instruction)
        )))

        resp = client.models.generate_content(model=model, contents=contents)
        parsed = _extract_json(getattr(resp, "text", "") or "")
        return {
            "success": True,
            "source": "frames_dir",
            "frames_dir": frames_dir,
            "analyzed_images": len([c for c in contents if getattr(c, "inline_data", None)]) or len(entries),
            "model": model,
            **parsed,
        }
    except Exception as e:
        return {"success": False, "error": f"Gemini解析エラー: {e}"}
