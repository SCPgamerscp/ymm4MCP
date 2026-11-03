"""Regression tests for extremely large speech-rate integers."""
import importlib.util
from pathlib import Path

import pytest


MODULE_PATH = Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py"
SPEC = importlib.util.spec_from_file_location("ymm4_editing", MODULE_PATH)
editing = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(editing)


@pytest.mark.parametrize("speed", [10 ** 400, -(10 ** 400), 0, True, float("inf"), float("nan")])
def test_plan_script_rejects_invalid_speech_rate_as_value_error(speed):
    with pytest.raises(ValueError, match="chars_per_sec"):
        editing.plan_script({"lines": [{"character": "霊夢", "text": "こんにちは"}],
                             "chars_per_sec": speed})


@pytest.mark.parametrize("speed", [1, 5.0, 1000])
def test_plan_script_accepts_valid_speech_rate(speed):
    result = editing.plan_script({"lines": [{"character": "霊夢", "text": "こんにちは"}],
                                  "chars_per_sec": speed})
    assert result["success"] is True
    assert result["dry_run"] is True
    assert result["added"] == 0
