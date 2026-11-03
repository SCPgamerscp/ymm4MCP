"""Reject invalid Unicode in planned script content before it reaches host JSON."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from editing import plan_script


@pytest.mark.parametrize("field", ["text", "character"])
@pytest.mark.parametrize("surrogate", ["\ud800", "\udfff"])
def test_script_plan_rejects_unpaired_surrogates(field, surrogate):
    line = {"text": "こんにちは", "character": "霊夢"}
    line[field] += surrogate
    with pytest.raises(ValueError, match="invalid Unicode surrogate"):
        plan_script({"lines": [line]})


def test_script_plan_accepts_valid_non_bmp_unicode():
    result = plan_script({"lines": [{"text": "絵文字😀", "character": "霊夢🌸"}]})
    assert result["success"] is True
    assert result["details"][0]["text"] == "絵文字😀"
