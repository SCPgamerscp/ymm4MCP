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
        description="YMM4を操作・情報取得するための単一ツール。action='get_info', 'control', 'add_item', 'edit_item', 'add_script'を指定して各種操作を行う。",
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
                    "description": "情報取得(status,project,items,effects_list)、操作(play,stop,save)、アイテム追加(text,voice,tachie,face)、編集(face_param,property,effect,delete,duration)のいずれか"
                },
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
                case _: raise ValueError(f"Unknown sub_action for get_info: {sub_action}")

        case "control":
            match sub_action:
                case "play": return await ymm4_post("/playback/play")
                case "stop": return await ymm4_post("/playback/stop")
                case "save": return await ymm4_post("/project/save")
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
                case _: raise ValueError(f"Unknown sub_action for edit_item: {sub_action}")

        case "add_script":
            return await add_script(args)

        case _:
            raise ValueError(f"Unknown action: {action}")


async def add_script(args: dict) -> dict:
    """
    台本をまとめてタイムラインに追加する
    キャラクターのセリフ数に応じて自動でフレームを割り当てる
    """
    lines = args.get("lines", [])
    fps = args.get("fps", 30)
    chars_per_sec = args.get("chars_per_sec", 5)
    current_frame = args.get("start_frame", 0)

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

        # 尺を文字数から自動計算 (最低1秒)
        estimated_secs = max(1.0, len(text) / chars_per_sec)
        length = int(estimated_secs * fps)

        res = await ymm4_post("/items/voice", {
            "text":      text,
            "character": character,
            "frame":     current_frame,
            "layer":     layer,
        })
        results.append({"character": character, "text": text[:20] + "...", "frame": current_frame, **res})
        current_frame += length

    return {"success": True, "added": len(results), "total_frames": current_frame, "details": results}


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
