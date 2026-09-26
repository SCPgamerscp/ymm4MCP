#!/usr/bin/env python3
"""
YMM4 MCP Server
YMM4(ゆっくりMovieMaker4)をMCP経由でClaudeから操作するサーバー

必要なもの:
  1. YMM4にMcpPluginをインストールして起動
  2. このサーバーをClaudeのMCP設定に追加

使い方 (claude_desktop_config.json):
  {
    "mcpServers": {
      "ymm4": {
        "command": "python",
        "args": ["C:/path/to/mcp-server/server.py"]
      }
    }
  }
"""

import asyncio
import json
import os
import sys
from typing import Any
from urllib.parse import quote
import httpx
from ymm4_connection import connection_settings, advanced_enabled
from editing import MAX_FRAME, integer, plan_script, validate_timeline, evaluate_qa_gate, finite_number
from jobs import job_id_ok, is_absolute_media_path, validate_export_request, validate_project_path
import editplan
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
    ImageContent,
    CallToolResult,
    ListToolsResult,
)

# Gemini 動画解析モジュール(同フォルダ)
try:
    import gemini_video
except Exception:
    # server.py が別CWDから起動された場合に備えてパスを通す
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import gemini_video  # type: ignore

# YMM4プラグインのHTTP API URL
YMM4_API_BASE = "http://127.0.0.1:8765/api"

from mcp_skills import register_skills

app = Server("ymm4-mcp")
register_skills(app)

# ============================================================
# HTTPクライアント
# ============================================================

HTTP_TIMEOUT_SECONDS = 10.0
PREVIEW_OVERHEAD_SECONDS = 15.0
_http_client: httpx.AsyncClient | None = None


def _get_http_client() -> httpx.AsyncClient:
    """同じイベントループ上でHTTP接続プールを再利用する。"""
    global _http_client
    if _http_client is None or _http_client.is_closed:
        _http_client = httpx.AsyncClient(timeout=HTTP_TIMEOUT_SECONDS, trust_env=False)
    return _http_client


async def close_http_client() -> None:
    """MCPサーバー終了時（別イベントループで再利用する前も）に接続を解放する。"""
    global _http_client
    client, _http_client = _http_client, None
    if client is not None:
        await client.aclose()


async def ymm4_get(path: str) -> dict:
    """YMM4 API GETリクエスト"""
    base, headers = connection_settings()
    res = await _get_http_client().get(f"{base}{path}", headers=headers)
    res.raise_for_status()
    return res.json()


async def ymm4_post(
    path: str, body: dict | None = None, *, timeout: float = HTTP_TIMEOUT_SECONDS
) -> dict:
    """応答待ち時間だけをリクエスト単位で変更し、接続待ちは10秒に保つ。"""
    base, headers = connection_settings()
    res = await _get_http_client().post(
        f"{base}{path}", headers=headers,
        json=body if body is not None else {},
        timeout=httpx.Timeout(HTTP_TIMEOUT_SECONDS, read=timeout),
    )
    res.raise_for_status()
    return res.json()


def _preview_timeout(duration_ms: int, minimum_ms: int) -> float:
    """C#側の録音時間クランプに合わせ、シーク・画像処理用の余裕を加える。"""
    if isinstance(duration_ms, bool) or not isinstance(duration_ms, int):
        raise ValueError("duration_ms は整数で指定してください")
    effective_ms = max(minimum_ms, min(duration_ms, 30000))
    return effective_ms / 1000.0 + PREVIEW_OVERHEAD_SECONDS

# ============================================================
# ツール定義
# ============================================================

TOOLS = [
    Tool(
        name="ymm4_interact",
        description=(
            "validate=タイムライン整合性・期待する配置の検証。qa_gate=修正ループの合格・継続・停止判定。add_scriptはdry_runで実行前に確認できます。"
            "plan_edit/apply_edit/reconcile_editで完成状態のEditPlanを差分適用できます。"
            "シーン失敗時は追加分だけrollbackし、完了済みシーンは残します。"
            "YMM4を操作・情報取得するための単一ツール。制作前にymm4://skills/{jikkyou,kaisetsu,chaban,story}の該当リソースを読んでください。"
            "action='get_info'(status/project/items/media/characters/capabilities/effects_list/effect_metadata/selection/commands/effects/keyframes/jobs/job/edit_state/checkpoints), "
            "'control'(play/stop/save/open/save_as/export/cancel_job/resume_job/checkpoint/rollback/undo/redo/split/align), "
            "'add_item'(video/audio/image/text/voice/tachie/face), "
            "'edit_item'(face_param/property/effect/delete/duration/move/select/resolve_overlaps/shift/keyframe), "
            "'add_script'(複数セリフ一括追加・実音声長で重なり自動回避), "
            "'plan_edit'(宣言的編集のdry-run), 'apply_edit'(差分適用・シーン単位rollback), 'reconcile_edit'(不足分だけ再実行)を指定する。"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["get_info", "control", "add_item", "edit_item", "add_script", "validate", "qa_gate",
                             "plan_edit", "apply_edit", "reconcile_edit"],
                    "description": "実行するアクションの種類"
                },
                "sub_action": {
                    "type": "string",
                    "description": (
                        "情報取得(status,project,items,media,characters,capabilities,effects_list,effect_metadata,selection,commands,effects,keyframes,jobs,job,edit_state,checkpoints)、"
                        "操作(play,stop,save,open,save_as,export,cancel_job,resume_job,checkpoint,rollback,undo,redo,split,align)、"
                        "アイテム追加(video,audio,image,text,voice,tachie,face)、"
                        "編集(face_param,property,effect,delete,duration,move,select,resolve_overlaps,shift,keyframe)のいずれか"
                    )
                },
                "dry_run": {"type": "boolean", "description": "add_script/apply_edit: 検証と推定配置のみ。編集・音声合成なし"},
                "plan": {"type": "object", "description": "plan_edit/apply_edit/reconcile_edit: 完成状態のEditPlan（scenesとitems）"},
                "atomic_scenes": {"type": "boolean", "description": "apply_edit: シーン途中の失敗でそのシーンの追加分を削除する。既定true"},
                "checkpoint_id": {"type": "string", "description": "get_info/checkpoints と control/rollback の対象"},
                "reason": {"type": "string", "description": "control/checkpoint: 操作ログに残す理由"},
                "backup": {"type": "boolean", "description": "control/checkpoint: 保存済み.ymmpがあればコピーする"},
                "expected": {"type": "array", "maxItems": 1000, "items": {"type": "object"}, "description": "validate: 配置後に一意に存在すべきitem_id/revision/frame/layer/length/type/text"},
                "duration": {"type": "integer", "minimum": 1, "description": "validate: プロジェクトの上限フレーム（省略可）"},
                "include_gaps": {"type": "boolean", "default": True, "description": "validate: 同一レイヤー内のアイテム間の空白を警告する"},
                "subtitle_layers": {"type": "array", "minItems": 1, "maxItems": 128,
                                    "items": {"type": "integer", "minimum": 0},
                                    "description": "validate: 指定レイヤーのTextItemを字幕として扱い、各VoiceItemとの時間・本文一致を検査（省略時は検査しない）"},
                "qa_history": {"type": "array", "maxItems": 20, "items": {"type": "object"},
                               "description": "qa_gate: 同じ検査条件で得た過去のvalidate結果。古い順"},
                "max_repairs": {"type": "integer", "minimum": 0, "maximum": 20, "description": "qa_gate: 最大修正回数。既定3"},
                "repeat_limit": {"type": "integer", "minimum": 2, "maximum": 10, "description": "qa_gate: 同じ問題群が連続したら停止。既定2"},
                "elapsed_seconds": {"type": "number", "minimum": 0, "description": "qa_gate: 呼び出し側で計測した修正ループ経過秒数"},
                "max_seconds": {"type": "number", "exclusiveMinimum": 0, "description": "qa_gate: 経過時間の上限（省略時は無制限）"},
                "api_calls": {"type": "integer", "minimum": 0, "description": "qa_gate: 呼び出し側で数えたAPI回数"},
                "max_api_calls": {"type": "integer", "minimum": 1, "description": "qa_gate: API回数の上限（省略時は無制限）"},
                "min_silence_seconds": {"type": "number", "minimum": 0.1, "maximum": 60, "description": "get_info/audio_qa: 無音とみなす最短秒数。既定2"},
                "path": {"type": "string", "description": "get_info/media/audio_qa、video/audio/imageの素材、またはexport/open/save_asの絶対パス"},
                "output_path": {"type": "string", "description": "export: 書き出し先の絶対パス（pathの別名）"},
                "format": {"type": "string", "enum": ["mp4", "wav", "avi", "mov", "mkv", "webm"], "description": "export: 出力形式。省略時は拡張子"},
                "overwrite": {"type": "boolean", "description": "export/save_as: 既存ファイルを上書きする"},
                "force": {"type": "boolean", "description": "open: 未保存変更または保存状態不明でもプロジェクトを開く"},
                "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 7200, "description": "export: 完了待ちの上限秒"},
                "idempotency_key": {"type": "string", "description": "export/apply_edit: 同じキーの再送は既存ジョブまたは既存適用を返す"},
                "job_id": {"type": "string", "description": "get_info/job と cancel_job/resume_job の対象"},
                "from_frame": {"type": "integer", "minimum": 0, "maximum": 2147483647, "description": "shift: このフレーム以降を対象"},
                "delta": {"type": "integer", "minimum": -2147483647, "maximum": 2147483647, "description": "shift: 加算するフレーム数(負で前詰め)。移動後の配置が範囲外なら全件拒否"},
                "gap": {"type": "integer", "minimum": 0, "maximum": 2147483647, "description": "resolve_overlaps: アイテム間の最小すき間フレーム"},
                "filename": {"type": "string", "minLength": 1, "description": "move: 対象アイテムのファイル名(部分一致、空白のみ不可)"},
                "clear": {"type": "boolean", "description": "select: trueで全選択解除"},
                "text": {"type": "string", "description": "表示または発話テキスト"},
                "character": {"type": "string", "description": "キャラクター名"},
                "frame": {"type": "integer", "minimum": -1, "maximum": 2147483647, "description": "通常は0以上。edit_item/deleteの-1だけ省略指定として許可"},
                "layer": {"type": "integer", "minimum": -1, "maximum": 2147483647, "description": "通常は0以上。edit_item/deleteの-1だけ省略指定として許可"},
                "length": {"type": "integer", "minimum": 1, "maximum": 2147483647},
                "item_id": {"type": "string", "description": "items/add_itemで返された安定ID。property/delete/select/keyframeではframe+layerより優先"},
                "expected_revision": {"type": "string", "description": "property/delete/keyframe時の楽観ロック。最新itemsのrevisionと不一致なら変更しない"},
                "prop": {"type": "string", "description": "property/keyframe: X, Y, Opacity, Zoom などのプロパティ名"},
                "value": {"description": "propertyでは文字列。keyframeでは数値"},
                "at": {"type": "integer", "minimum": 0, "description": "keyframe: アイテム開始からの相対フレーム"},
                "keyframe_action": {"type": "string", "enum": ["set", "remove", "clear"], "description": "keyframe: set=打刻, remove=1点削除, clear=全削除"},
                "effect": {"type": "string"},
                "name": {"type": "string", "description": "effect_metadata: effects_list の name または fullName"},
                "params": {"type": "object"},
                "frames": {"type": "integer", "minimum": 1, "maximum": 2147483647},
                "layers": {"type": "array", "maxItems": 1000, "items": {"type": "integer", "minimum": 0, "maximum": 2147483647}},
                "lines": {
                    "type": "array",
                    "items": {"type": "object", "properties": {"character": {"type": "string"}, "text": {"type": "string"}, "layer": {"type": "integer"}}}
                },
                "fps": {"type": "integer", "default": 30},
                "chars_per_sec": {"type": "number", "default": 5},
                "start_frame": {"type": "integer", "default": 0}
            },
            "required": ["action"]
        }
    ),
    Tool(
        name="ymm4_preview",
        description="YMM4のプレビュー映像・音声を認識するツール。capture=現在フレームを画像取得、seek_capture=指定フレームに移動して画像取得、position=再生位置取得、record=システム音声録音(ループバック)",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["capture", "seek_capture", "position", "record", "watch"],
                    "description": "capture:現在画像取得 / seek_capture:フレーム移動+画像 / position:再生位置 / record:音声録音 / watch:映像+音声同時取得"
                },
                "frame": {"type": "integer", "description": "seek_capture/watchで移動するフレーム番号"},
                "duration_ms": {"type": "integer", "description": "record/watchで録音する時間(ms)、デフォルト3000/5000"},
                "capture_interval_ms": {"type": "integer", "description": "watchでフレームをキャプチャする間隔(ms)、デフォルト1000"},
                "element": {"type": "string", "description": "キャプチャ対象のWPF要素名(省略可)"}
            },
            "required": ["action"]
        }
    ),
    Tool(
        name="ymm4_advanced",
        description=(
            "【YMM4全機能アクセス】個別ツールで未対応のYMM4内部機能に、リフレクション経由で直接アクセスする上級ツール。"
            "YMM4内部の任意のViewModel/Model/Projectのプロパティ取得・設定、任意メソッド呼び出し、任意コマンド実行、"
            "オブジェクト構造の調査(inspect)ができる。"
            "\n\naction一覧:\n"
            "- inspect: 対象オブジェクトのプロパティ・メソッド・コマンド一覧を取得（まず構造を調べる時に使う）\n"
            "- get: 任意プロパティ/フィールドの現在値を取得（path指定で深掘り可: 'ActiveTimeline.Items[0].Item.Length'）\n"
            "- set: 任意プロパティ/フィールドに値を設定\n"
            "- invoke: 任意メソッドを引数付きで呼び出す（戻り値がTaskなら自動await）\n"
            "- command: 任意のICommandを実行（UndoCommand等UIメニュー限定機能を直接トリガー）\n"
            "- list_commands: 利用可能な全コマンドと現在実行可能かを一覧\n"
            "\ntarget一覧: 'Main'(MainViewModel) / 'ActiveTimeline' / 'Player' / 'Project'。"
            "pathで 'A.B[2].C' のようにドット・インデックスで深掘りできる。ReactivePropertyの.Valueは自動展開される。"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["inspect", "get", "set", "invoke", "command", "list_commands"],
                    "description": "inspect:構造調査 / get:値取得 / set:値設定 / invoke:メソッド呼出 / command:コマンド実行 / list_commands:コマンド一覧"
                },
                "target": {
                    "type": "string",
                    "description": "対象オブジェクト: Main / ActiveTimeline / Player / Project（省略時Main）。任意のプロパティ名も可。"
                },
                "path": {
                    "type": "string",
                    "description": "get/set/inspectでの深掘りパス。例: 'Items[0].Item.Length' や 'Project.Name'（省略可）"
                },
                "name": {"type": "string", "description": "command/list_commandsでのコマンド名（例: UndoCommand）"},
                "method": {"type": "string", "description": "invokeで呼び出すメソッド名"},
                "args": {"type": "array", "description": "invokeのメソッド引数（順序通り）", "items": {}},
                "param": {"description": "commandの実行パラメータ（省略可）"},
                "value": {"description": "setで設定する値"}
            },
            "required": ["action"]
        }
    ),
    Tool(
        name="ymm4_analyze_video",
        description=(
            "【Gemini動画解析】動画を時系列で精密に解析し、何が起きているか(シーン変化/キャラ登場/効果音/"
            "テロップ/動き等)をタイムスタンプ付きで検出する。検出したいイベントは event_instruction で自由にカスタマイズ可能"
            "(未指定なら全イベントを検出)。\n\n"
            "2つのモードがある:\n"
            "- source='preview': YMM4プレビューの指定フレーム区間を連番画像+音声として書き出し、Geminiで解析する。"
            "タイムラインに配置済みの素材の内容を解析したい時に使う。start_frame/end_frame/step_frames を指定。\n"
            "- source='file': 取り込み前の元動画ファイル(mp4等)をGeminiに直接渡して解析する。video_path を指定。\n\n"
            "結果のイベントには time_sec(動画先頭からの秒) と、それを変換した frame(YMM4フレーム番号) が含まれるため、"
            "そのまま add_script のセリフ配置やイベント同期に使える。\n"
            "※ 環境変数 GEMINI_API_KEY が必要。"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "source": {
                    "type": "string",
                    "enum": ["preview", "file"],
                    "description": "preview:YMM4プレビュー区間を解析 / file:元動画ファイルを解析"
                },
                "event_instruction": {
                    "type": "string",
                    "description": "検出したいイベントの指示(日本語可)。例:'効果音とキャラの登場だけ検出して'。省略時は全イベント検出。"
                },
                "video_path": {"type": "string", "description": "source='file'時: 解析する動画ファイルの絶対パス"},
                "start_frame": {"type": "integer", "description": "source='preview'時: 解析開始フレーム(既定0)"},
                "end_frame": {"type": "integer", "description": "source='preview'時: 解析終了フレーム"},
                "step_frames": {"type": "integer", "description": "source='preview'時: 何フレームおきにキャプチャするか(既定3。細かくするほど精密だが遅い)"},
                "record_audio": {"type": "boolean", "description": "source='preview'時: 区間音声も録音して音響イベント解析に使うか(既定true)"},
                "base_frame": {"type": "integer", "description": "file解析時に time_sec→frame 変換の基準にする開始フレーム(既定0)。検出結果をこのフレーム以降に配置したい時に使う。"},
                "fps": {"type": "integer", "description": "file解析時のFPS(省略時はYMM4から取得、取得不可なら30)"},
                "model": {"type": "string", "description": "使用するGeminiモデル(既定 gemini-2.5-flash。高精度には gemini-2.5-pro)"}
            },
            "required": ["source"]
        }
    )
]

# ============================================================
# ハンドラー
# ============================================================

@app.list_tools()
async def list_tools() -> ListToolsResult:
    return ListToolsResult(tools=[tool for tool in TOOLS if tool.name != "ymm4_advanced" or advanced_enabled()])


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> CallToolResult:
    try:
        if name == "ymm4_preview":
            result = await dispatch_preview(arguments)
            return result
        if name == "ymm4_advanced":
            if not advanced_enabled():
                raise ValueError("高度APIは無効です。YMM4_ENABLE_ADVANCED=1とプラグイン側の許可が必要です")
            result = await dispatch_advanced(arguments)
            return CallToolResult(content=[TextContent(type="text", text=format_result(result))],
                                  isError=isinstance(result, dict) and (result.get("success") is False or "error" in result))
        if name == "ymm4_analyze_video":
            result = await analyze_video(arguments)
            return CallToolResult(content=[TextContent(type="text", text=format_result(result))],
                                  isError=isinstance(result, dict) and (result.get("success") is False or "error" in result))
        if name != "ymm4_interact":
            raise ValueError(f"Unknown tool: {name}")
        result = await dispatch(arguments)
        return CallToolResult(content=[TextContent(type="text", text=format_result(result))],
                                  isError=isinstance(result, dict) and (result.get("success") is False or "error" in result))
    except (httpx.ConnectError, httpx.ConnectTimeout):
        msg = (
            "❌ YMM4プラグインサーバーに接続できません。\n"
            "YMM4を起動し、ツールメニューから「MCP連携サーバー」を開いて「▶ 起動」ボタンを押してください。"
        )
        return CallToolResult(content=[TextContent(type="text", text=msg)], isError=True)
    except httpx.HTTPStatusError as exc:
        status = exc.response.status_code
        try:
            error = exc.response.json()
        except ValueError:
            error = {}
        if isinstance(error, dict) and error.get("error_code"):
            return CallToolResult(content=[TextContent(type="text", text=format_result(error))], isError=True)
        message = {401: "認証失敗。プラグインを再起動し接続情報を確認してください", 403: "APIの利用が許可されていません"}.get(status, f"YMM4 HTTPエラー: {status}")
        return CallToolResult(content=[TextContent(type="text", text=message)], isError=True)
    except httpx.TimeoutException:
        msg = (
            "YMM4プラグインサーバーの処理待ちがタイムアウトしました。\n"
            "録音時間を短くするか、YMM4の処理完了を待って状態を確認してください。"
            "タイムアウト後も処理が続いている可能性があるため、自動再試行はしていません。"
        )
        return CallToolResult(content=[TextContent(type="text", text=msg)], isError=True)
    except Exception as e:
        return CallToolResult(
            content=[TextContent(type="text", text=f"❌ エラー: {str(e)}")],
            isError=True,
        )


def require_edit_target(args: dict, *, item_id_allowed: bool = True, partial: bool = False) -> None:
    """Reject edit requests that would silently target the item at frame 0, layer 0."""
    if "item_id" in args:
        if not item_id_allowed:
            raise ValueError("this edit action requires frame and layer, not item_id")
        if not isinstance(args["item_id"], str) or not args["item_id"].strip():
            raise ValueError("item_id must be a non-empty string")
        return
    if partial:
        if "frame" not in args and "layer" not in args:
            raise ValueError("select requires item_id, frame, layer, or clear=true")
    elif "frame" not in args or "layer" not in args:
        raise ValueError("frame and layer are required when item_id is not specified")


async def dispatch(args: dict) -> Any:
    action = args.get("action")
    sub_action = args.get("sub_action")
    
    match action:
        case "get_info":
            match sub_action:
                case "status": return await ymm4_get("/status")
                case "characters": return await ymm4_get("/characters")
                case "capabilities": return await ymm4_get("/capabilities")
                case "project": return await ymm4_get("/project")
                case "items": return await ymm4_get("/items")
                case "audio_qa":
                    path = args.get("path")
                    if not is_absolute_media_path(path):
                        raise ValueError("audio_qa path must be an absolute file path")
                    q = f"/media/audio-qa?path={quote(path.strip(), safe='')}"
                    if "min_silence_seconds" in args:
                        seconds = finite_number(args["min_silence_seconds"], "min_silence_seconds")
                        if not 0.1 <= seconds <= 60:
                            raise ValueError("min_silence_seconds must be in 0.1..60")
                        q += f"&min_silence_seconds={seconds:g}"
                    return await ymm4_get(q)
                case "media":
                    path = args.get("path")
                    if not is_absolute_media_path(path):
                        raise ValueError("media path must be an absolute file path")
                    return await ymm4_get(f"/media/info?path={quote(path.strip(), safe='')}")
                case "effects_list": return await ymm4_get("/effects/list")
                case "effect_metadata":
                    name = args.get("name")
                    if not isinstance(name, str) or not name.strip():
                        raise ValueError("effect_metadata requires name")
                    return await ymm4_get(f"/effects/describe?name={quote(name.strip(), safe='')}")
                case "selection": return await ymm4_get("/selection")
                case "commands": return await ymm4_get("/commands")
                case "position": return await ymm4_get("/preview/position")
                case "effects":
                    q = []
                    if "item_id" in args: q.append(f"item_id={quote(str(args['item_id']), safe='')}")
                    if "frame" in args: q.append(f"frame={args['frame']}")
                    if "layer" in args: q.append(f"layer={args['layer']}")
                    qs = ("?" + "&".join(q)) if q else ""
                    return await ymm4_get(f"/items/effects{qs}")
                case "keyframes":
                    q = []
                    if "item_id" in args: q.append(f"item_id={quote(str(args['item_id']), safe='')}")
                    if "frame" in args: q.append(f"frame={args['frame']}")
                    if "layer" in args: q.append(f"layer={args['layer']}")
                    if "prop" in args: q.append(f"prop={quote(str(args['prop']), safe='')}")
                    qs = ("?" + "&".join(q)) if q else ""
                    return await ymm4_get(f"/items/keyframes{qs}")
                case "jobs": return await ymm4_get("/jobs")
                case "job":
                    job_id = args.get("job_id")
                    if not job_id_ok(job_id):
                        raise ValueError("job_id is invalid")
                    return await ymm4_get(f"/jobs/{job_id}")
                case "edit_state": return await ymm4_get("/edits/state")
                case "checkpoints":
                    checkpoint_id = args.get("checkpoint_id")
                    if checkpoint_id is None:
                        return await ymm4_get("/edits/checkpoints")
                    if not editplan.checkpoint_id_ok(checkpoint_id):
                        raise ValueError("checkpoint_id is invalid")
                    return await ymm4_get(f"/edits/checkpoints/{checkpoint_id}")
                case _: raise ValueError(f"Unknown sub_action for get_info: {sub_action}")

        case "control":
            match sub_action:
                case "play": return await ymm4_post("/playback/play")
                case "stop": return await ymm4_post("/playback/stop")
                case "save": return await ymm4_post("/project/save")
                case "open":
                    body = {"path": validate_project_path(args, must_exist=False)}
                    if "force" in args:
                        if not isinstance(args["force"], bool):
                            raise ValueError("force must be boolean")
                        body["force"] = args["force"]
                    return await ymm4_post("/project/open", body)
                case "save_as":
                    body = {"path": validate_project_path(args, must_exist=False)}
                    if "overwrite" in args:
                        if not isinstance(args["overwrite"], bool):
                            raise ValueError("overwrite must be boolean")
                        body["overwrite"] = args["overwrite"]
                    return await ymm4_post("/project/save-as", body)
                case "export":
                    parsed = validate_export_request(args)
                    body = {
                        "path": parsed["output_path"],
                        "format": parsed["format"],
                        "overwrite": parsed["overwrite"],
                        "timeout_seconds": parsed["timeout_seconds"],
                    }
                    if parsed["idempotency_key"]:
                        body["idempotency_key"] = parsed["idempotency_key"]
                    return await ymm4_post("/project/export", body)
                case "cancel_job":
                    job_id = args.get("job_id")
                    if not job_id_ok(job_id):
                        raise ValueError("job_id is invalid")
                    return await ymm4_post(f"/jobs/{job_id}/cancel")
                case "resume_job":
                    job_id = args.get("job_id")
                    if not job_id_ok(job_id):
                        raise ValueError("job_id is invalid")
                    return await ymm4_post(f"/jobs/{job_id}/resume")
                case "checkpoint":
                    body = {}
                    reason = editplan.parse_reason(args)
                    if reason is not None:
                        body["reason"] = reason
                    if "backup" in args:
                        if not isinstance(args["backup"], bool):
                            raise ValueError("backup must be boolean")
                        body["backup"] = args["backup"]
                    return await ymm4_post("/edits/checkpoint", body)
                case "rollback":
                    checkpoint_id = args.get("checkpoint_id")
                    if not editplan.checkpoint_id_ok(checkpoint_id):
                        raise ValueError("checkpoint_id is invalid")
                    return await ymm4_post("/edits/rollback", {"checkpoint_id": checkpoint_id})
                # YMM4内部コマンドをトリガー（UIメニュー限定機能を直接実行）
                case "undo": return await ymm4_post("/command", {"name": "UndoCommand"})
                case "redo": return await ymm4_post("/command", {"name": "RedoCommand"})
                case "split": return await ymm4_post("/command", {"name": "SplitItemCommand", "target": "ActiveTimeline"})
                case "align": return await ymm4_post("/command", {"name": "AlignItemsCommand", "target": "ActiveTimeline"})
                case _: raise ValueError(f"Unknown sub_action for control: {sub_action}")

        case "add_item":
            payload = {}
            if "text" in args: payload["text"] = args["text"]
            if "character" in args: payload["character"] = args["character"]
            if "frame" in args: payload["frame"] = args["frame"]
            if "layer" in args: payload["layer"] = args["layer"]
            if "length" in args: payload["length"] = args["length"]
            if "path" in args: payload["path"] = args["path"]
            for key in ("frame", "layer"):
                if key in payload: integer(payload[key], key)
            if "length" in payload: integer(payload["length"], "length", 1)
            if sub_action == "image" and "length" not in payload:
                raise ValueError("image requires length")
            if sub_action in ("video", "audio", "image") and not isinstance(payload.get("path"), str):
                raise ValueError("media requires path")

            match sub_action:
                case "video" | "audio" | "image": return await ymm4_post(f"/items/{sub_action}", payload, timeout=120.0)
                case "text": return await ymm4_post("/items/text", payload)
                case "voice": return await ymm4_post("/items/voice", payload, timeout=120.0)
                case "tachie": return await ymm4_post("/items/tachie", payload)
                case "face": return await ymm4_post("/items/face", payload)
                case _: raise ValueError(f"Unknown sub_action for add_item: {sub_action}")

        case "edit_item":
            # Reject malformed coordinates before a mutating request reaches YMM4.  Delete's
            # historical -1 selector means "not specified", so it remains supported there.
            selector_minimum = -1 if sub_action == "delete" else 0
            for key in ("frame", "layer"):
                if key in args:
                    integer(args[key], key, selector_minimum)
            if "layers" in args:
                if not isinstance(args["layers"], list) or len(args["layers"]) > 1000:
                    raise ValueError("layers must be an array of at most 1000 layer numbers")
                for index, layer in enumerate(args["layers"]):
                    integer(layer, f"layers[{index}]")

            match sub_action:
                case "face_param":
                    require_edit_target(args, item_id_allowed=False)
                    payload = args.get("params", {})
                    if "frame" in args: payload["frame"] = args["frame"]
                    if "layer" in args: payload["layer"] = args["layer"]
                    return await ymm4_post("/items/face/param", payload)
                case "property":
                    if "expected_revision" in args and not args.get("item_id"):
                        raise ValueError("expected_revision を使う場合は item_id も指定してください")
                    require_edit_target(args)
                    payload = {
                        "frame": args.get("frame", 0),
                        "layer": args.get("layer", 0),
                        "prop": args.get("prop", ""),
                        "value": str(args.get("value", "")),
                    }
                    if "item_id" in args: payload["item_id"] = args["item_id"]
                    if "expected_revision" in args: payload["expected_revision"] = args["expected_revision"]
                    return await ymm4_post("/items/prop", payload)
                case "effect":
                    require_edit_target(args, item_id_allowed=False)
                    return await ymm4_post("/items/effect", {
                        "frame": args.get("frame", 0), 
                        "layer": args.get("layer", 0), 
                        "effect": args.get("effect", "")
                    })
                case "delete":
                    if "expected_revision" in args and not args.get("item_id"):
                        raise ValueError("expected_revision を使う場合は item_id も指定してください")
                    payload = {}
                    if "item_id" in args: payload["item_id"] = args["item_id"]
                    if "expected_revision" in args: payload["expected_revision"] = args["expected_revision"]
                    if "frame" in args and args["frame"] != -1: payload["frame"] = args["frame"]
                    if "layer" in args and args["layer"] != -1: payload["layer"] = args["layer"]
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/items/delete", payload)
                case "duration":
                    return await ymm4_post("/timeline/duration", {
                        "frames": integer(args.get("frames", 0), "frames", 1),
                    })
                case "move":
                    if not isinstance(args.get("filename"), str) or not args["filename"].strip():
                        raise ValueError("move requires a non-empty filename")
                    payload = {"filename": args.get("filename", ""),
                               "frame": integer(args.get("frame", 0), "frame")}
                    if "length" in args:
                        payload["length"] = integer(args["length"], "length", 1)
                    return await ymm4_post("/items/move", payload)
                case "select":
                    if args.get("clear") is not True:
                        require_edit_target(args, partial=True)
                    payload = {}
                    if "item_id" in args: payload["item_id"] = args["item_id"]
                    if "frame" in args: payload["frame"] = args["frame"]
                    if "layer" in args: payload["layer"] = args["layer"]
                    if args.get("clear"): payload["clear"] = True
                    return await ymm4_post("/items/select", payload)
                case "resolve_overlaps":
                    payload = {}
                    if "gap" in args: payload["gap"] = integer(args["gap"], "gap")
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/timeline/resolve-overlaps", payload)
                case "shift":
                    payload = {
                        "fromFrame": integer(args.get("from_frame", 0), "from_frame"),
                        "delta": integer(args.get("delta", 0), "delta", -MAX_FRAME),
                    }
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/timeline/shift", payload)
                case "keyframe":
                    if "expected_revision" in args and not args.get("item_id"):
                        raise ValueError("expected_revision を使う場合は item_id も指定してください")
                    require_edit_target(args)
                    payload = {
                        "prop": args.get("prop", ""),
                        "action": args.get("keyframe_action", "set"),
                    }
                    if "item_id" in args: payload["item_id"] = args["item_id"]
                    if "expected_revision" in args: payload["expected_revision"] = args["expected_revision"]
                    if "frame" in args: payload["frame"] = args["frame"]
                    if "layer" in args: payload["layer"] = args["layer"]
                    if "at" in args: payload["at"] = integer(args["at"], "at")
                    elif "keyframe" in args: payload["at"] = integer(args["keyframe"], "keyframe")
                    if payload["action"] == "set":
                        if "value" not in args:
                            raise ValueError("keyframe set には value が必要です")
                        payload["value"] = finite_number(args["value"], "value")
                    return await ymm4_post("/items/keyframe", payload)
                case _: raise ValueError(f"Unknown sub_action for edit_item: {sub_action}")

        case "validate" | "qa_gate":
            snapshot = await ymm4_get("/items")
            if snapshot.get("success") is False or "error" in snapshot:
                return snapshot
            qa = validate_timeline(snapshot.get("items"), args.get("expected"), args.get("duration"),
                                   args.get("include_gaps", True), args.get("subtitle_layers"))
            if action == "validate":
                return qa
            return evaluate_qa_gate(
                qa, args.get("qa_history"), max_repairs=args.get("max_repairs", 3),
                repeat_limit=args.get("repeat_limit", 2), elapsed_seconds=args.get("elapsed_seconds", 0),
                max_seconds=args.get("max_seconds"), api_calls=args.get("api_calls", 0),
                max_api_calls=args.get("max_api_calls"))

        case "add_script":
            return await add_script(args)

        case "plan_edit":
            return await run_edit_plan(args, dry_run=True)

        case "apply_edit":
            return await run_edit_plan(args, dry_run=bool(args.get("dry_run", False)))

        case "reconcile_edit":
            return await run_edit_plan(args, dry_run=False)

        case _:
            raise ValueError(f"Unknown action: {action}")


async def add_script(args: dict) -> dict:
    """Validate the entire script first; never continue after failed/unknown synthesis."""
    plan = plan_script(args)
    if args.get("dry_run", False):
        return plan
    characters = await ymm4_get("/characters")
    if characters.get("success") is False or "error" in characters:
        return characters
    names = [c["name"] for c in characters.get("characters", [])]
    for line in plan["details"]:
        if names.count(line["character"]) != 1:
            raise ValueError(f"キャラ名は一覧から一意の完全一致名を指定してください: {line['character']}")
    frame = args.get("start_frame", 0)
    gap = args.get("gap", 0)
    results = []
    for index, line in enumerate(plan["details"]):
        try:
            res = await ymm4_post("/items/voice", {
                "text": line["text"], "character": line["character"],
                "frame": frame, "layer": line["layer"],
            }, timeout=120.0)
        except httpx.HTTPError:
            return {"success": False, "error_code": "SCRIPT_REQUEST_FAILED",
                    "error": "通信失敗。追加された可能性があるためitemsで確認してください。自動再試行なし。",
                    "outcome_unknown": True, "added": len(results), "failed_line": index,
                    "details": results, "rolled_back": False}
        if not isinstance(res, dict) or res.get("success") is not True:
            return {"success": False, "error_code": "SCRIPT_PARTIAL_FAILURE",
                    "error": "セリフ追加に失敗したため停止しました", "failed_line": index,
                    "added": len(results), "details": results, "failure": res, "rolled_back": False}
        results.append(res)
        try:
            length = integer(res.get("length"), "actual voice length", 1)
            actual_frame = integer(res.get("frame"), "actual frame")
            frame = integer(actual_frame + length + gap, "next frame")
        except ValueError:
            return {"success": False, "error_code": "VOICE_LENGTH_UNKNOWN",
                    "error": "追加結果の実長を確定できません。推定尺で続行せずitemsで確認してください。",
                    "added": len(results), "failed_line": index, "details": results,
                    "rolled_back": False}
    try:
        snapshot = await ymm4_get("/items")
    except httpx.HTTPError as exc:
        snapshot = {"error": str(exc)}
    if (not isinstance(snapshot, dict) or snapshot.get("success") is False
            or "error" in snapshot or not isinstance(snapshot.get("items"), list)):
        return {"success": False, "error_code": "SCRIPT_VERIFY_FAILED",
                "error": "追加後のタイムラインを確認できません。itemsを確認してから再実行してください",
                "added": len(results), "details": results, "verification": snapshot,
                "outcome_unknown": True, "rolled_back": False}
    details = [{"id": str(index), "op": "add", "item_id": res.get("item_id"), "result": res}
               for index, res in enumerate(results)]
    failures = editplan.verify_applied_items(details, snapshot["items"])
    if failures:
        return {"success": False, "error_code": "SCRIPT_VERIFY_FAILED",
                "error": "追加結果と現在のタイムラインが一致しません。itemsを確認してください",
                "added": len(results), "details": results, "verification_failures": failures,
                "outcome_unknown": True, "rolled_back": False}
    return {"success": True, "added": len(results), "total_frames": frame,
            "details": results, "verified": True, "dry_run": False}


async def _load_edit_binding(key: str | None) -> dict | None:
    if not key:
        return None
    state = await ymm4_get("/edits/state")
    if (not isinstance(state, dict) or state.get("success") is False or "error" in state
            or not isinstance(state.get("bindings"), list)):
        raise ValueError("EditPlan bindings are unavailable")
    return editplan.pick_binding(state.get("bindings"), key)


async def _delete_item_ids(item_ids: list[str]) -> dict:
    ids = [item_id for item_id in item_ids if isinstance(item_id, str) and item_id]
    if not ids:
        return {"success": True, "removed": 0, "item_ids": []}
    try:
        res = await ymm4_post("/items/delete", {"item_ids": ids})
    except httpx.HTTPError:
        return {"success": False, "error_code": "ROLLBACK_REQUEST_FAILED",
                "error": "ロールバック削除の通信に失敗しました。itemsで確認してください。",
                "outcome_unknown": True, "item_ids": ids}
    if not isinstance(res, dict) or res.get("success") is not True:
        return {"success": False, "error_code": "ROLLBACK_FAILED",
                "error": "失敗シーンの追加分を削除できませんでした",
                "failure": res, "item_ids": ids}
    return {"success": True, "removed": res.get("removed", len(ids)),
            "item_ids": res.get("item_ids", ids), "missing": res.get("missing") or []}


async def _persist_edit_bindings(plan: dict, bindings: list[dict]) -> dict | None:
    if not plan.get("idempotency_key") or not bindings:
        return None
    try:
        saved = await ymm4_post("/edits/bindings", {
            "idempotency_key": plan["idempotency_key"],
            "plan_hash": editplan.plan_hash(plan),
            "items": bindings,
        })
        if isinstance(saved, dict) and saved.get("success") is False:
            return saved
    except httpx.HTTPError:
        return {"error_code": "BINDINGS_NOT_SAVED"}
    return None


async def _partial_edit_result(plan, *, error_code, error, failed_id, failed_index, failed_scene,
                               added, details, bindings, committed, scenes, scene_added_ids,
                               atomic, scene_details=None, scene_bindings=None, extra=None,
                               outcome_unknown=False):
    rollback = None
    rolled_back = False
    remaining_added = added
    remaining_details = list(details)
    if atomic and scene_added_ids and not outcome_unknown:
        rollback = await _delete_item_ids(scene_added_ids)
        rolled_back = rollback.get("success") is True
        if rolled_back:
            remaining_added = max(0, added - len(scene_added_ids))
        elif scene_details:
            remaining_details = remaining_details + scene_details
    elif scene_details:
        remaining_details = remaining_details + scene_details
    persist = list(bindings)
    if not rolled_back and scene_bindings:
        persist.extend(scene_bindings)
    binding_error = await _persist_edit_bindings(plan, persist)
    scenes = list(scenes)
    scenes.append({
        "id": failed_scene, "status": "rolled_back" if rolled_back else "partial",
        "added": 0 if rolled_back else len(scene_added_ids),
        "failed_id": failed_id,
    })
    result = {
        "success": False, "error_code": error_code, "error": error,
        "failed_id": failed_id, "failed_index": failed_index, "failed_scene": failed_scene,
        "added": remaining_added, "details": remaining_details, "committed_scenes": committed,
        "scenes": scenes, "rolled_back": rolled_back, "atomic_scenes": atomic,
        "plan_hash": editplan.plan_hash(plan), "idempotency_key": plan.get("idempotency_key"),
        "item_count": plan["item_count"],
        "note": ("Failed scene additions were removed; committed scenes were kept."
                 if rolled_back else
                 "Failed scene was not rolled back. Inspect items, then reconcile_edit."),
    }
    if extra:
        result.update(extra)
    if outcome_unknown:
        result["outcome_unknown"] = True
    if rollback is not None:
        result["rollback"] = rollback
    if binding_error:
        result["warning"] = "BINDINGS_NOT_SAVED"
        result["binding_error"] = binding_error
    return result


async def run_edit_plan(args: dict, dry_run: bool) -> dict:
    """Validate an EditPlan, optionally replay an idempotent apply, otherwise add only missing items."""
    plan = editplan.parse_plan(args)
    characters = await ymm4_get("/characters")
    names = None
    if isinstance(characters, dict) and characters.get("success") is not False and "error" not in characters:
        names = [c["name"] for c in characters.get("characters", []) if isinstance(c, dict) and "name" in c]
    snapshot = await ymm4_get("/items")
    if not isinstance(snapshot, dict):
        return {"success": False, "error_code": "ITEMS_UNAVAILABLE",
                "error": "タイムラインの状態を取得できません。編集せず停止しました"}
    if snapshot.get("success") is False or "error" in snapshot:
        return snapshot
    items = snapshot.get("items")
    if not isinstance(items, list):
        return {"success": False, "error_code": "ITEMS_UNAVAILABLE",
                "error": "タイムラインの状態を取得できません。編集せず停止しました"}
    try:
        record = await _load_edit_binding(plan.get("idempotency_key"))
    except (httpx.HTTPError, ValueError):
        return {"success": False, "error_code": "BINDINGS_UNAVAILABLE",
                "error": "再適用情報を確認できません。編集せず停止しました", "rolled_back": False}
    conflict = editplan.binding_conflict(plan, items, record)
    if conflict:
        return conflict
    if not dry_run:
        replay = editplan.replay_record(plan, items, record)
        if replay:
            return replay
    preview = editplan.diff_plan(plan, items, record, names)
    if dry_run:
        if names is None:
            preview.setdefault("warnings", []).append({
                "code": "CHARACTERS_UNAVAILABLE", "severity": "warning",
                "message": characters.get("error") if isinstance(characters, dict) else "characters unavailable",
            })
        preview["dry_run"] = True
        return preview
    if names is None:
        return characters if isinstance(characters, dict) else {
            "success": False, "error_code": "CHARACTERS_UNAVAILABLE", "error": "キャラ一覧を取得できません"}
    unknown = next((w for w in preview["warnings"] if w.get("code") == "CHARACTER_UNKNOWN"), None)
    if unknown:
        raise ValueError(f"キャラ名は一覧から一意の完全一致名を指定してください: {unknown.get('character')}")
    content_conflicts = [w for w in preview["warnings"] if w.get("code") == "CONTENT_STATE_CONFLICT"]
    if content_conflicts:
        return {"success": False, "error_code": "EDIT_STATE_CONFLICT",
                "error": "同じ内容の既存アイテムの配置が計画と異なります",
                "conflicts": content_conflicts, "rolled_back": False}

    current_by_id = {item.get("item_id"): item for item in items if isinstance(item, dict)}
    gap = plan["gap"]
    cursor = {scene["id"]: scene["start_frame"] for scene in plan["scenes"]}
    atomic = bool(plan.get("atomic_scenes", True))
    bindings = []
    details = []
    committed = []
    scene_reports = []
    added = 0
    op_index = -1
    for scene in editplan.group_ops_by_scene(preview["ops"]):
        scene_id = scene["id"]
        scene_added_ids = []
        scene_bindings = []
        scene_details = []
        scene_added = 0
        scene_kept = 0
        for op in scene["ops"]:
            op_index += 1
            if op["op"] == "keep":
                current = current_by_id.get(op.get("item_id"), {})
                try:
                    end = integer(current.get("frame", op["frame"]), "frame") + integer(
                        current.get("length", op["length"]), "length", 1)
                    cursor[scene_id] = max(cursor.get(scene_id, 0), end + gap)
                except ValueError:
                    pass
                row = {"id": op["id"], "op": "keep", "item_id": op.get("item_id"),
                       "revision": op.get("revision"), "scene_id": scene_id}
                scene_bindings.append({"id": op["id"], "item_id": op.get("item_id"), "revision": op.get("revision")})
                scene_details.append(row)
                scene_kept += 1
                continue
            payload = dict(op["payload"])
            if op.get("frame_source") == "sequential":
                payload["frame"] = cursor.get(scene_id, payload.get("frame", 0))
            try:
                res = await ymm4_post(f"/items/{op['kind']}", payload, timeout=120.0)
            except httpx.HTTPError:
                return await _partial_edit_result(
                    plan, error_code="EDIT_REQUEST_FAILED",
                    error="通信失敗。追加された可能性があるためitemsで確認してください。自動再試行なし。",
                    failed_id=op["id"], failed_index=op_index, failed_scene=scene_id,
                    added=added, details=details, bindings=bindings, committed=committed,
                    scenes=scene_reports, scene_added_ids=scene_added_ids, atomic=atomic,
                    scene_details=scene_details, scene_bindings=scene_bindings, outcome_unknown=True)
            if not isinstance(res, dict) or res.get("success") is not True:
                return await _partial_edit_result(
                    plan, error_code="EDIT_PARTIAL_FAILURE",
                    error="EditPlanの適用に失敗したため停止しました",
                    failed_id=op["id"], failed_index=op_index, failed_scene=scene_id,
                    added=added, details=details, bindings=bindings, committed=committed,
                    scenes=scene_reports, scene_added_ids=scene_added_ids, atomic=atomic,
                    scene_details=scene_details, scene_bindings=scene_bindings, extra={"failure": res})
            added += 1
            scene_added += 1
            item_id = res.get("item_id")
            if isinstance(item_id, str) and item_id:
                scene_added_ids.append(item_id)
            scene_details.append({"id": op["id"], "op": "add", "item_id": item_id,
                                  "revision": res.get("revision"), "kind": op["kind"],
                                  "scene_id": scene_id, "result": res})
            scene_bindings.append({"id": op["id"], "item_id": item_id, "revision": res.get("revision")})
            try:
                length = integer(res.get("length"), "actual length", 1)
                actual_frame = integer(res.get("frame"), "actual frame")
                cursor[scene_id] = integer(actual_frame + length + gap, "next frame")
            except ValueError:
                if op["kind"] == "voice":
                    return await _partial_edit_result(
                        plan, error_code="VOICE_LENGTH_UNKNOWN",
                        error="追加結果の実長を確定できません。推定尺で続行せずitemsで確認してください。",
                        failed_id=op["id"], failed_index=op_index, failed_scene=scene_id,
                        added=added, details=details, bindings=bindings, committed=committed,
                        scenes=scene_reports, scene_added_ids=scene_added_ids, atomic=atomic,
                        scene_details=scene_details, scene_bindings=scene_bindings)
                cursor[scene_id] = payload.get("frame", 0) + op.get("length", 1) + gap
        bindings.extend(scene_bindings)
        details.extend(scene_details)
        committed.append(scene_id)
        scene_reports.append({
            "id": scene_id, "status": "committed", "added": scene_added, "kept": scene_kept,
        })

    try:
        verified_snapshot = await ymm4_get("/items")
    except httpx.HTTPError as exc:
        verified_snapshot = {"error": str(exc)}
    if (not isinstance(verified_snapshot, dict) or verified_snapshot.get("success") is False
            or "error" in verified_snapshot or not isinstance(verified_snapshot.get("items"), list)):
        return {"success": False, "error_code": "EDIT_VERIFY_FAILED",
                "error": "適用後のタイムラインを確認できません。itemsを確認してから再実行してください",
                "added": added, "details": details, "outcome_unknown": True, "rolled_back": False,
                "atomic_scenes": atomic, "committed_scenes": committed, "scenes": scene_reports,
                "verification": verified_snapshot}
    verification_failures = editplan.verify_applied_items(details, verified_snapshot["items"])
    if verification_failures:
        return {"success": False, "error_code": "EDIT_VERIFY_FAILED",
                "error": "追加結果と現在のタイムラインが一致しません。itemsを確認してください",
                "added": added, "details": details, "verification_failures": verification_failures,
                "outcome_unknown": True, "rolled_back": False, "atomic_scenes": atomic,
                "committed_scenes": committed, "scenes": scene_reports}

    if plan.get("idempotency_key"):
        try:
            saved = await ymm4_post("/edits/bindings", {
                "idempotency_key": plan["idempotency_key"],
                "plan_hash": editplan.plan_hash(plan),
                "items": bindings,
            })
        except httpx.HTTPError as exc:
            saved = {"error": str(exc)}
        if not isinstance(saved, dict) or saved.get("success") is not True:
            code = saved.get("error_code") if isinstance(saved, dict) else None
            return {"success": False, "error_code": code or "BINDINGS_NOT_SAVED",
                    "error": "EditPlanは適用されましたが、再適用情報を保存できませんでした",
                    "added": added, "details": details, "binding_error": saved,
                    "rolled_back": False, "atomic_scenes": atomic,
                    "committed_scenes": committed, "scenes": scene_reports}
    result = {
        "success": True, "replayed": False, "dry_run": False, "added": added,
        "kept": sum(d["op"] == "keep" for d in details), "details": details,
        "plan_hash": editplan.plan_hash(plan), "idempotency_key": plan.get("idempotency_key"),
        "total_frames": max(cursor.values()) if cursor else plan["total_frames"],
        "item_count": plan["item_count"], "rolled_back": False, "atomic_scenes": atomic,
        "committed_scenes": committed, "scenes": scene_reports,
        "note": "Each scene is a transaction: a failed scene rolls back its own additions only.",
    }
    return result


async def dispatch_advanced(args: dict) -> Any:
    """
    全機能アクセス用の上級ツール。
    YMM4内部の任意のオブジェクトに対してリフレクション経由でアクセスする。
    """
    action = args.get("action")
    target = args.get("target", "Main")

    match action:
        case "inspect":
            q = [f"target={quote(target, safe='')}"]
            if args.get("path"):
                q.append(f"path={quote(args['path'], safe='')}")
            return await ymm4_get("/reflect/inspect?" + "&".join(q))

        case "get":
            return await ymm4_post("/reflect/get", {
                "target": target,
                "path": args.get("path", ""),
            })

        case "set":
            payload = {"target": target, "path": args.get("path", ""), "value": args.get("value")}
            return await ymm4_post("/reflect/set", payload)

        case "invoke":
            return await ymm4_post("/reflect/invoke", {
                "target": target,
                "method": args.get("method", ""),
                "args": args.get("args", []),
            })

        case "command":
            payload = {"name": args.get("name", ""), "target": target}
            if "param" in args:
                payload["param"] = args["param"]
            return await ymm4_post("/command", payload)

        case "list_commands":
            return await ymm4_get("/commands")

        case _:
            raise ValueError(f"Unknown action for ymm4_advanced: {action}")


# ============================================================
# Gemini 動画解析
# ============================================================

async def ymm4_post_long(path: str, body: dict, timeout: float = 600.0) -> dict:
    """長時間処理(クリップ書き出し等)用のPOST。タイムアウトを長めに取る。"""
    return await ymm4_post(path, body, timeout=timeout)


def _sec_to_frame(time_sec: Any, fps: int, base_frame: int) -> Any:
    """秒 → YMM4フレーム番号に変換する。"""
    try:
        return base_frame + int(round(float(time_sec) * fps))
    except (TypeError, ValueError):
        return None


def _attach_frames(result: dict, fps: int, base_frame: int) -> dict:
    """Geminiが返したeventsの time_sec/end_sec をフレーム番号に変換して付加する。"""
    events = result.get("events")
    if isinstance(events, list):
        for ev in events:
            if not isinstance(ev, dict):
                continue
            if "time_sec" in ev:
                ev["frame"] = _sec_to_frame(ev.get("time_sec"), fps, base_frame)
            if "end_sec" in ev and ev.get("end_sec") is not None:
                ev["end_frame"] = _sec_to_frame(ev.get("end_sec"), fps, base_frame)
    result["fps_used"] = fps
    result["base_frame"] = base_frame
    return result


async def _get_fps_from_ymm4(default: int = 30) -> int:
    """YMM4プラグインからFPSを取得する。取れなければdefault。"""
    try:
        data = await ymm4_get("/project/fps")
        fps = data.get("fps")
        if isinstance(fps, int) and fps > 0:
            return fps
    except Exception:
        pass
    return default


async def analyze_video(args: dict) -> dict:
    """
    動画をGeminiで解析し、タイムスタンプ+YMM4フレーム番号付きのイベントを返す。

    source='preview': C#の /api/preview/export-clip で区間を連番PNG+WAVに書き出し、
                      gemini_video.analyze_frames_dir で解析する。
    source='file'   : gemini_video.analyze_video_file で元動画を直接解析する。
    """
    # Geminiが使えるか先にチェック
    ok, msg = gemini_video.is_available()
    if not ok:
        return {
            "success": False,
            "error": msg,
            "hint": "Google AI Studioでキーを発行し、Claude Desktop設定のenvで GEMINI_API_KEY を渡すか、"
                    "サーバー起動環境で環境変数を設定してください。`pip install google-genai` も必要です。",
        }

    source = args.get("source")
    instruction = args.get("event_instruction")  # Noneなら全イベント
    model = args.get("model") or gemini_video.DEFAULT_MODEL

    if source == "file":
        video_path = args.get("video_path", "")
        if not video_path:
            return {"success": False, "error": "source='file'にはvideo_pathが必要です"}
        base_frame = int(args.get("base_frame", 0))
        fps = args.get("fps")
        if not isinstance(fps, int) or fps <= 0:
            fps = await _get_fps_from_ymm4(30)

        # Gemini呼び出しはブロッキングなので別スレッドで
        result = await asyncio.to_thread(
            gemini_video.analyze_video_file, video_path, instruction, model
        )
        if result.get("success"):
            result = _attach_frames(result, fps, base_frame)
        return result

    elif source == "preview":
        start_frame = int(args.get("start_frame", 0))
        end_frame = int(args.get("end_frame", start_frame + 300))
        step_frames = int(args.get("step_frames", 3))
        record_audio = args.get("record_audio", True)

        # 1) C#でプレビュー区間を連番PNG+WAVに書き出す(長時間)
        try:
            clip = await ymm4_post_long("/preview/export-clip", {
                "startFrame": start_frame,
                "endFrame": end_frame,
                "stepFrames": step_frames,
                "recordAudio": record_audio,
            })
        except (httpx.ConnectError, httpx.TimeoutException):
            raise
        except Exception as e:
            return {"success": False, "error": f"クリップ書き出し失敗: {e}"}

        if not clip.get("success"):
            return {"success": False, "error": clip.get("error", "export-clip失敗"), "detail": clip}

        frames_dir = clip.get("outputDir")
        frames_meta = clip.get("frames", [])
        audio_path = clip.get("audioFile")
        fps = clip.get("fps", 30)

        # 2) GeminiでフレームフォルダをN解析(ブロッキング→別スレッド)
        result = await asyncio.to_thread(
            gemini_video.analyze_frames_dir,
            frames_dir, frames_meta, audio_path, instruction, model,
        )
        if result.get("success"):
            # time_sec はクリップ先頭=0 基準なので、base_frame=start_frame で実フレームに戻す
            result = _attach_frames(result, fps, start_frame)
            result["clip"] = {
                "outputDir": frames_dir,
                "frameCount": clip.get("frameCount"),
                "audioFile": audio_path,
                "hasAudio": clip.get("hasAudio"),
            }
        return result

    else:
        return {"success": False, "error": f"不明なsource: {source} (preview または file)"}


def format_result(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


async def dispatch_preview(args: dict) -> CallToolResult:
    """プレビュー系ツールのディスパッチ。画像はImageContentで返しClaudeが直接認識できる。"""
    action = args.get("action")

    match action:
        case "capture":
            params = ""
            if "element" in args:
                params = f"?element={args['element']}"
            data = await ymm4_get(f"/preview/capture{params}")
            return _preview_result(data)

        case "seek_capture":
            frame = args.get("frame", 0)
            data = await ymm4_post(
                "/preview/seek", {"frame": frame}, timeout=PREVIEW_OVERHEAD_SECONDS
            )
            return _preview_result(data)

        case "position":
            data = await ymm4_get("/preview/position")
            return CallToolResult(content=[TextContent(type="text", text=format_result(data))])

        case "record":
            duration_ms = args.get("duration_ms", 3000)
            data = await ymm4_post(
                "/preview/record", {"duration_ms": duration_ms},
                timeout=_preview_timeout(duration_ms, 500),
            )
            if data.get("success") is False or "error" in data:
                return CallToolResult(content=[TextContent(type="text", text=format_result(data))], isError=True)
            audio_b64 = data.pop("audio", None)
            summary = format_result(data)
            contents: list = [TextContent(type="text", text=summary)]
            if audio_b64:
                import tempfile, os, base64
                wav_bytes = base64.b64decode(audio_b64)
                # 固定パスに保存（上書き）してClaudeが参照しやすくする
                save_dir = os.path.dirname(os.path.abspath(__file__))
                wav_path = os.path.join(save_dir, "..", "record_result.wav")
                wav_path = os.path.normpath(wav_path)
                with open(wav_path, "wb") as f:
                    f.write(wav_bytes)
                has_audio = data.get("has_audio", False)
                rms = data.get("rms_level", 0)
                contents.append(TextContent(
                    type="text",
                    text=(
                        f"\n🎵 録音完了: {wav_path}\n"
                        f"   RMSレベル: {rms} ({'音声あり ✅' if has_audio else '無音またはほぼ無音 ⚠️'})\n"
                        f"   サイズ: {len(wav_bytes):,} bytes"
                    )
                ))
            return CallToolResult(content=contents)

        case "watch":
            # 映像＋音声を同時取得
            frame = args.get("frame", 0)
            duration_ms = args.get("duration_ms", 5000)
            interval_ms = args.get("capture_interval_ms", 1000)
            data = await ymm4_post("/preview/watch", {
                "frame": frame,
                "duration_ms": duration_ms,
                "capture_interval_ms": interval_ms
            }, timeout=_preview_timeout(duration_ms, 1000))
            if not data.get("success"):
                return CallToolResult(
                    content=[TextContent(type="text", text=f"❌ {data.get('error', 'unknown error')}")],
                    isError=True
                )
            contents: list = []
            # 音声情報
            audio = data.get("audio", {})
            rms = audio.get("rms_level", 0)
            has_audio = audio.get("has_audio", False)
            contents.append(TextContent(
                type="text",
                text=(
                    f"🎬 watch: frame={data.get('start_frame')} duration={data.get('duration_ms')}ms\n"
                    f"🎵 音声: RMS={rms} {'✅' if has_audio else '⚠️無音'}"
                )
            ))
            # 音声WAVを保存
            audio_b64 = audio.get("data")
            if audio_b64:
                import base64, os
                wav_bytes = base64.b64decode(audio_b64)
                save_dir = os.path.dirname(os.path.abspath(__file__))
                wav_path = os.path.normpath(os.path.join(save_dir, "..", "watch_result.wav"))
                with open(wav_path, "wb") as f:
                    f.write(wav_bytes)
                contents.append(TextContent(type="text", text=f"💾 音声保存: {wav_path}"))
            # 各フレーム画像
            for frame_data in data.get("frames", []):
                img_b64 = frame_data.get("image")
                t = frame_data.get("time_ms", 0)
                if img_b64:
                    contents.append(TextContent(type="text", text=f"📸 t={t}ms"))
                    contents.append(ImageContent(
                        type="image",
                        data=img_b64,
                        mimeType="image/png"
                    ))
            return CallToolResult(content=contents)

        case _:
            raise ValueError(f"Unknown preview action: {action}")


def _preview_result(data: dict) -> CallToolResult:
    """C#から返った {success, image(base64 PNG), ...} をImageContentに変換"""
    if not data.get("success", False):
        return CallToolResult(
            content=[TextContent(type="text", text=f"❌ {data.get('error', 'unknown error')}")],
            isError=True
        )
    image_b64 = data.get("image")
    if not image_b64:
        return CallToolResult(
            content=[TextContent(type="text", text=format_result(data))],
            isError=True
        )
    meta = {k: v for k, v in data.items() if k != "image"}
    return CallToolResult(content=[
        ImageContent(type="image", data=image_b64, mimeType="image/png"),
        TextContent(type="text", text=f"📸 {meta}")
    ])


# ============================================================
# エントリーポイント
# ============================================================

async def main():
    try:
        async with stdio_server() as (read_stream, write_stream):
            await app.run(read_stream, write_stream, app.create_initialization_options())
    finally:
        await close_http_client()


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="YMM4 MCP server")
    parser.add_argument("--transport", choices=("stdio", "streamable-http"), default="stdio")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8766)
    parser.add_argument("--tls-cert")
    parser.add_argument("--tls-key")
    options = parser.parse_args()
    if options.transport == "stdio":
        asyncio.run(main())
    else:
        from http_transport import run_http
        asyncio.run(run_http(app, close_http_client, options))
