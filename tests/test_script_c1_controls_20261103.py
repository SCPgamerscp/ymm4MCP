"""Regression: reject C1 dialogue controls while allowing multiline text."""
import importlib.util
from pathlib import Path
import pytest

spec = importlib.util.spec_from_file_location("editing", Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py")
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)

@pytest.mark.parametrize("control", ["\x80", "\x85", "\x9f"])
def test_reject_c1_dialogue_controls(control):
    with pytest.raises(ValueError, match="text contains a control character"):
        editing.plan_script({"lines": [{"character": "霊夢", "text": "hello" + control}]})

def test_preserve_multiline_dialogue():
    result = editing.plan_script({"lines": [{"character": "霊夢", "text": "こんにちは\n世界\t！\r"}]})
    assert result["details"][0]["text"] == "こんにちは\n世界\t！\r"

@pytest.mark.parametrize("control", [
    "\u061c", "\u200e", "\u200f", "\u202a", "\u202b", "\u202c",
    "\u202d", "\u202e", "\u2066", "\u2067", "\u2068", "\u2069",
])
def test_reject_bidi_controls_in_dialogue_text(control):
    with pytest.raises(ValueError, match="text contains a control character"):
        editing.plan_script({"lines": [{"character": "霊夢", "text": "安全" + control + "な台本"}]})

def test_preserve_normal_international_dialogue():
    text = "日本語とالعربيةとEnglish\n字幕\tです"
    result = editing.plan_script({"lines": [{"character": "霊夢", "text": text}]})
    assert result["details"][0]["text"] == text
