"""Regression tests for bounded character identifiers in script planning."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from editing import plan_script


def test_character_name_over_limit_is_rejected():
    with pytest.raises(ValueError, match=r"lines\[0\]\.character exceeds 256"):
        plan_script({"lines": [{"text": "hello", "character": "A" * 257}]})


def test_character_name_at_limit_is_accepted():
    result = plan_script({"lines": [{"text": "hello", "character": "A" * 256}]})
    assert result["success"] is True
    assert result["details"][0]["character"] == "A" * 256


def test_long_dialogue_text_still_uses_its_own_limit():
    result = plan_script({"lines": [{"text": "x" * 10000, "character": "Reimu"}]})
    assert result["details"][0]["text"] == "x" * 10000


@pytest.mark.parametrize("bad_name", [" Reimu", "Reimu ", "Reimu\n", "Reimu\x00", "Reimu\x80"])
def test_ambiguous_character_name_is_rejected(bad_name):
    with pytest.raises(ValueError, match="character (must not have surrounding whitespace|contains a control character)"):
        plan_script({"lines": [{"text": "hello", "character": bad_name}]})


def test_internal_space_and_japanese_name_remain_valid():
    result = plan_script({"lines": [{"text": "hello", "character": "博麗 霊夢"}]})
    assert result["details"][0]["character"] == "博麗 霊夢"
