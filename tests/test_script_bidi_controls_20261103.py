"""Regression tests for Unicode bidi spoofing in speaker identifiers."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from editing import plan_script


@pytest.mark.parametrize("control", [
    "\u061c", "\u200e", "\u200f", "\u202a", "\u202b",
    "\u202c", "\u202d", "\u202e", "\u2066", "\u2067",
    "\u2068", "\u2069",
])
def test_bidi_controls_are_not_allowed_in_character_names(control):
    with pytest.raises(ValueError, match="control character"):
        plan_script({"lines": [{"text": "hello", "character": "霊夢" + control}]})


def test_normal_unicode_character_name_remains_valid():
    result = plan_script({"lines": [{"text": "こんにちは", "character": "博麗 霊夢"}]})
    assert result["success"] is True
    assert result["details"][0]["character"] == "博麗 霊夢"
