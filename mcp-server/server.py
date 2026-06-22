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
import httpx
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
YMM4_API_BASE = "http://localhost:8765/api"

app = Server("ymm4-mcp")

# ============================================================
# HTTPクライアント
# ============================================================

async def ymm4_get(path: str) -> dict:
    """YMM4 API GETリクエスト"""
    async with httpx.AsyncClient(timeout=10.0) as client:
        res = await client.get(f"{YMM4_API_BASE}{path}")
        res.raise_for_status()
        return res.json()

async def ymm4_post(path: str, body: dict = {}) -> dict:
    """YMM4 API POSTリクエスト"""
    async with httpx.AsyncClient(timeout=10.0) as client:
        res = await client.post(f"{YMM4_API_BASE}{path}", json=body)
        res.raise_for_status()
        return res.json()

# ============================================================
# ツール定義
# ============================================================

TOOLS = [
    Tool(
        name="ymm4_interact",
        description=(
            "YMM4を操作・情報取得するための単一ツール。"
            "action='get_info'(status/project/items/effects_list/selection/commands/effects), "
            "'control'(play/stop/save/undo/redo/split/align), "
            "'add_item'(text/voice/tachie/face), "
            "'edit_item'(face_param/property/effect/delete/duration/move/select/resolve_overlaps/shift), "
            "'add_script'(複数セリフ一括追加・実音声長で重なり自動回避)を指定する。"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["get_info", "control", "add_item", "edit_item", "add_script"],
                    "description": "実行するアクションの種類"
                },
                "sub_action": {
                    "type": "string",
                    "description": (
                        "情報取得(status,project,items,effects_list,selection,commands,effects)、"
                        "操作(play,stop,save,undo,redo,split,align)、"
                        "アイテム追加(text,voice,tachie,face)、"
                        "編集(face_param,property,effect,delete,duration,move,select,resolve_overlaps,shift)のいずれか"
                    )
                },
                "from_frame": {"type": "integer", "description": "shift: このフレーム以降を対象"},
                "delta": {"type": "integer", "description": "shift: 加算するフレーム数(負で前詰め)"},
                "gap": {"type": "integer", "description": "resolve_overlaps: アイテム間の最小すき間フレーム"},
                "filename": {"type": "string", "description": "move: 対象アイテムのファイル名(部分一致)"},
                "clear": {"type": "boolean", "description": "select: trueで全選択解除"},
                "text": {"type": "string", "description": "表示または発話テキスト"},
                "character": {"type": "string", "description": "キャラクター名"},
                "frame": {"type": "integer"},
                "layer": {"type": "integer"},
                "length": {"type": "integer"},
                "prop": {"type": "string"},
                "value": {"type": "string"},
                "effect": {"type": "string"},
                "params": {"type": "object"},
                "frames": {"type": "integer"},
                "layers": {"type": "array", "items": {"type": "integer"}},
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
    return ListToolsResult(tools=TOOLS)


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> CallToolResult:
    try:
        if name == "ymm4_preview":
            result = await dispatch_preview(arguments)
            return result
        if name == "ymm4_advanced":
            result = await dispatch_advanced(arguments)
            return CallToolResult(content=[TextContent(type="text", text=format_result(result))])
        if name == "ymm4_analyze_video":
            result = await analyze_video(arguments)
            return CallToolResult(content=[TextContent(type="text", text=format_result(result))])
        if name != "ymm4_interact":
            raise ValueError(f"Unknown tool: {name}")
        result = await dispatch(arguments)
        return CallToolResult(content=[TextContent(type="text", text=format_result(result))])
    except httpx.ConnectError:
        msg = (
            "❌ YMM4プラグインサーバーに接続できません。\n"
            "YMM4を起動し、ツールメニューから「MCP連携サーバー」を開いて「▶ 起動」ボタンを押してください。"
        )
        return CallToolResult(content=[TextContent(type="text", text=msg)], isError=True)
    except Exception as e:
        return CallToolResult(
            content=[TextContent(type="text", text=f"❌ エラー: {str(e)}")],
            isError=True,
        )


async def dispatch(args: dict) -> Any:
    action = args.get("action")
    sub_action = args.get("sub_action")
    
    match action:
        case "get_info":
            match sub_action:
                case "status": return await ymm4_get("/status")
                case "project": return await ymm4_get("/project")
                case "items": return await ymm4_get("/items")
                case "effects_list": return await ymm4_get("/effects/list")
                case "selection": return await ymm4_get("/selection")
                case "commands": return await ymm4_get("/commands")
                case "position": return await ymm4_get("/preview/position")
                case "effects":
                    q = []
                    if "frame" in args: q.append(f"frame={args['frame']}")
                    if "layer" in args: q.append(f"layer={args['layer']}")
                    qs = ("?" + "&".join(q)) if q else ""
                    return await ymm4_get(f"/items/effects{qs}")
                case _: raise ValueError(f"Unknown sub_action for get_info: {sub_action}")

        case "control":
            match sub_action:
                case "play": return await ymm4_post("/playback/play")
                case "stop": return await ymm4_post("/playback/stop")
                case "save": return await ymm4_post("/project/save")
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
            
            match sub_action:
                case "text": return await ymm4_post("/items/text", payload)
                case "voice": return await ymm4_post("/items/voice", payload)
                case "tachie": return await ymm4_post("/items/tachie", payload)
                case "face": return await ymm4_post("/items/face", payload)
                case _: raise ValueError(f"Unknown sub_action for add_item: {sub_action}")

        case "edit_item":
            match sub_action:
                case "face_param":
                    payload = args.get("params", {})
                    if "frame" in args: payload["frame"] = args["frame"]
                    if "layer" in args: payload["layer"] = args["layer"]
                    return await ymm4_post("/items/face/param", payload)
                case "property":
                    return await ymm4_post("/items/prop", {
                        "frame": args.get("frame", 0), 
                        "layer": args.get("layer", 0), 
                        "prop": args.get("prop", ""), 
                        "value": str(args.get("value", ""))
                    })
                case "effect":
                    return await ymm4_post("/items/effect", {
                        "frame": args.get("frame", 0), 
                        "layer": args.get("layer", 0), 
                        "effect": args.get("effect", "")
                    })
                case "delete":
                    payload = {}
                    if "frame" in args and args["frame"] != -1: payload["frame"] = args["frame"]
                    if "layer" in args and args["layer"] != -1: payload["layer"] = args["layer"]
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/items/delete", payload)
                case "duration":
                    return await ymm4_post("/timeline/duration", {"frames": args.get("frames", 0)})
                case "move":
                    payload = {"filename": args.get("filename", ""), "frame": args.get("frame", 0)}
                    if "length" in args: payload["length"] = args["length"]
                    return await ymm4_post("/items/move", payload)
                case "select":
                    payload = {}
                    if "frame" in args: payload["frame"] = args["frame"]
                    if "layer" in args: payload["layer"] = args["layer"]
                    if args.get("clear"): payload["clear"] = True
                    return await ymm4_post("/items/select", payload)
                case "resolve_overlaps":
                    payload = {}
                    if "gap" in args: payload["gap"] = args["gap"]
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/timeline/resolve-overlaps", payload)
                case "shift":
                    payload = {"fromFrame": args.get("from_frame", 0), "delta": args.get("delta", 0)}
                    if "layers" in args: payload["layers"] = args["layers"]
                    return await ymm4_post("/timeline/shift", payload)
                case _: raise ValueError(f"Unknown sub_action for edit_item: {sub_action}")

        case "add_script":
            return await add_script(args)

        case _:
            raise ValueError(f"Unknown action: {action}")


async def add_script(args: dict) -> dict:
    """
    台本をまとめてタイムラインに追加する。

    重なり防止の核心:
      各セリフを1件追加するごとに、C#側が返す「実際の音声長(length/フレーム数)」を使って
      次のセリフの開始フレームを動的に決定する。これにより文字数推定のズレによる
      アイテムの重なりを根本的に防止する。
      C#が実長を取得できなかった場合(length<=0)のみ、文字数からの推定値にフォールバックする。

    gap(フレーム)を指定すると各セリフ間にすき間を空ける。
    """
    lines = args.get("lines", [])
    fps = args.get("fps", 30)
    chars_per_sec = args.get("chars_per_sec", 5)
    current_frame = args.get("start_frame", 0)
    gap = args.get("gap", 0)

    # キャラクターごとのデフォルトレイヤー
    char_layer_map: dict[str, int] = {}
    next_layer = 0

    results = []
    for line in lines:
        character = line.get("character", "ゆっくり霊夢")
        text = line.get("text", "")
        layer = line.get("layer")

        # レイヤーが未指定ならキャラクターに自動割り当て
        if layer is None:
            if character not in char_layer_map:
                char_layer_map[character] = next_layer
                next_layer += 1
            layer = char_layer_map[character]

        # 文字数からの推定尺 (最低1秒) — 実長が取れない場合のフォールバック
        estimated_secs = max(1.0, len(text) / chars_per_sec)
        estimated_length = int(estimated_secs * fps)

        res = await ymm4_post("/items/voice", {
            "text":      text,
            "character": character,
            "frame":     current_frame,
            "layer":     layer,
        })

        # C#が返した実音声長を優先。取れなければ推定値を使う。
        actual_length = res.get("length", -1) if isinstance(res, dict) else -1
        used_length = actual_length if isinstance(actual_length, int) and actual_length > 0 else estimated_length
        length_source = "actual" if used_length == actual_length and actual_length > 0 else "estimated"

        results.append({
            "character": character,
            "text": (text[:20] + "...") if len(text) > 20 else text,
            "frame": current_frame,
            "length": used_length,
            "length_source": length_source,
            **(res if isinstance(res, dict) else {"raw": res}),
        })
        current_frame += used_length + gap

    return {
        "success": True,
        "added": len(results),
        "total_frames": current_frame,   # 次シーンのstart_frameの目安
        "details": results,
    }


async def dispatch_advanced(args: dict) -> Any:
    """
    全機能アクセス用の上級ツール。
    YMM4内部の任意のオブジェクトに対してリフレクション経由でアクセスする。
    """
    action = args.get("action")
    target = args.get("target", "Main")

    match action:
        case "inspect":
            q = [f"target={target}"]
            if args.get("path"):
                q.append(f"path={args['path']}")
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
    async with httpx.AsyncClient(timeout=timeout) as client:
        res = await client.post(f"{YMM4_API_BASE}{path}", json=body)
        res.raise_for_status()
        return res.json()


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
        except httpx.ConnectError:
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
            data = await ymm4_post("/preview/seek", {"frame": frame})
            return _preview_result(data)

        case "position":
            data = await ymm4_get("/preview/position")
            return CallToolResult(content=[TextContent(type="text", text=format_result(data))])

        case "record":
            duration_ms = args.get("duration_ms", 3000)
            data = await ymm4_post("/preview/record", {"duration_ms": duration_ms})
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
            })
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
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
