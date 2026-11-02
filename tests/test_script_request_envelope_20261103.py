"""Regression tests for malformed script-planning request envelopes."""
import importlib.util
from pathlib import Path

import pytest

MODULE = Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py"
spec = importlib.util.spec_from_file_location("editing_script_envelope", MODULE)
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)


@pytest.mark.parametrize("args", [None, [], "not-an-object", 1, True])
def test_plan_script_rejects_non_object_envelope(args):
    with pytest.raises(ValueError, match="script plan arguments must be an object"):
        editing.plan_script(args)


def test_plan_script_accepts_valid_object_envelope():
    result = editing.plan_script({"lines": [{"text": "こんにちは", "character": "霊夢"}]})
    assert result["success"] is True
    assert len(result["details"]) == 1
