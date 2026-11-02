"""Regression: script planning must never imply that edits were performed."""
import importlib.util
from pathlib import Path

import pytest

MODULE = Path(__file__).resolve().parents[1] / "mcp-server" / "editing.py"
spec = importlib.util.spec_from_file_location("editing_script_dry_run_regression", MODULE)
editing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editing)


def _args(**extra):
    return {"lines": [{"text": "こんにちは", "character": "霊夢"}], **extra}


@pytest.mark.parametrize("value", [False, None, 0, 1, "false", "true", [], {}])
def test_reject_non_true_dry_run(value):
    with pytest.raises(ValueError, match="dry_run must be true"):
        editing.plan_script(_args(dry_run=value))


@pytest.mark.parametrize("extra", [{}, {"dry_run": True}])
def test_script_plan_is_always_dry_run(extra):
    result = editing.plan_script(_args(**extra))
    assert result["success"] is True
    assert result["dry_run"] is True
    assert result["added"] == 0
    assert result["details"]
