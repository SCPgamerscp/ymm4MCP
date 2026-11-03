"""Regression: ducking must not trust visually spoofed host item IDs/revisions."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from ducking import valid_identifier, verify_keyframes


@pytest.mark.parametrize("control", [
    "\u061c", "\u200e", "\u200f", "\u202a", "\u202b", "\u202c",
    "\u202d", "\u202e", "\u2066", "\u2067", "\u2068", "\u2069",
])
def test_rejects_bidi_spoofing_in_identifiers(control):
    assert not valid_identifier("item" + control + "1")
    assert not valid_identifier("rev" + control + "1", max_length=256)


def test_preserves_valid_international_identifiers():
    assert valid_identifier("霊夢-العربية-😀")
    assert valid_identifier("revision-42", max_length=256)


def test_rejects_spoofed_revision_before_readback_acceptance():
    report = {"success": True, "revision": "rev\u202e1", "animations": []}
    assert not verify_keyframes(report, "rev\u202e1", [], 1.0)
