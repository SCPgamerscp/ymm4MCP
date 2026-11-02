"""Regression tests for unsafe dialogue control characters."""
import importlib.util
from pathlib import Path

import pytest

path = Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py"
spec = importlib.util.spec_from_file_location("editing_control_test", path)
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)


def plan(text):
    return editing.plan_script({"lines": [{"character": "霊夢", "text": text}]})


@pytest.mark.parametrize("text", ["a\x00b", "a\x01b", "a\x1fb", "a\x7fb"])
def test_rejects_unsafe_controls(text):
    with pytest.raises(ValueError, match="control character"):
        plan(text)


@pytest.mark.parametrize("text", ["こんにちは", "一行目\n二行目", "一\t二", "一\r\n二"])
def test_preserves_normal_and_multiline_dialogue(text):
    result = plan(text)
    assert result["success"] is True
    assert result["details"][0]["text"] == text
