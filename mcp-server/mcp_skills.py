"""Expose the existing, allowlisted skill files through standard MCP discovery."""
import json
from pathlib import Path
from mcp.types import Resource, Prompt, PromptArgument, GetPromptResult, PromptMessage, TextContent

SKILLS = {
    "jikkyou": "ゆっくり実況",
    "kaisetsu": "ゆっくり解説",
    "chaban": "ゆっくり茶番",
    "story": "ゆっくりストーリー",
}
ROOT = Path(__file__).resolve().parent / "skills"


def skill_text(name: str) -> str:
    if name not in SKILLS:
        raise ValueError(f"Unknown skill: {name}")
    return (ROOT / f"ymm4-{name}" / "SKILL.md").read_text(encoding="utf-8")


async def list_resources():
    return [Resource(uri=f"ymm4://skills/{name}", name=title, mimeType="text/markdown",
                     description=f"{title}の役割・話速・制作ワークフロー")
            for name, title in SKILLS.items()]


async def read_resource(uri):
    mapping = {f"ymm4://skills/{name}": name for name in SKILLS}
    name = mapping.get(str(uri))
    if name is None:
        raise ValueError(f"Unknown skill resource: {uri}")
    return skill_text(name)


async def list_prompts():
    return [Prompt(name=name, description=f"{title}の制作を開始する", arguments=[
        PromptArgument(name="theme", description="テーマ・制作の目的", required=True),
        PromptArgument(name="video_path", description="対象動画のローカルパス", required=False),
        PromptArgument(name="duration_seconds", description="目安の尺（正の整数秒）", required=False),
    ]) for name, title in SKILLS.items()]


async def get_prompt(name: str, arguments: dict | None = None):
    text = skill_text(name)
    args = arguments or {}
    if set(args) - {"theme", "video_path", "duration_seconds"}:
        raise ValueError("Unknown prompt argument")
    if not isinstance(args.get("theme"), str) or not args["theme"].strip():
        raise ValueError("theme は必須です")
    if "duration_seconds" in args:
        try:
            duration = int(args["duration_seconds"])
        except (ValueError, TypeError) as exc:
            raise ValueError("duration_seconds は正の整数秒で指定してください") from exc
        if duration <= 0 or str(duration) != str(args["duration_seconds"]):
            raise ValueError("duration_seconds は正の整数秒で指定してください")
    request = json.dumps(args, ensure_ascii=False)
    return GetPromptResult(description=SKILLS[name], messages=[PromptMessage(
        role="user", content=TextContent(type="text", text=(
            text + "\n\n## 今回の制作条件\n" + request +
            "\n最初にymm4_interact(action='get_info', sub_action='characters')でキャラ一覧を確認し、"
            "名前を完全一致で指定すること。台本をdry_runで確認してから配置し、"
            "配置後にitemsとvalidateで結果を検証すること。"
            "未対応の書き出しや未検証の結果を完了と報告しないこと。"
        ))
    )])


def register_skills(app):
    app.list_resources()(list_resources)
    app.read_resource()(read_resource)
    app.list_prompts()(list_prompts)
    app.get_prompt()(get_prompt)
