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

@pytest.mark.parametrize("control", [
    "\u200b", "\u200c", "\u200d", "\u2060", "\ufeff",
])
def test_rejects_invisible_format_characters_in_identifiers(control):
    assert not valid_identifier("item" + control + "1")
    assert not valid_identifier("rev" + control + "1", max_length=256)


def test_rejects_invisible_revision_before_readback_acceptance():
    report = {"success": True, "revision": "rev\u200b1", "animations": []}
    assert not verify_keyframes(report, "rev\u200b1", [], 1.0)


def test_expected_start_requires_absolute_frame_evidence():
    report = {"success": True, "revision": "rev1", "item_id": "bgm1",
              "animations": [{"prop": "Volume", "keyframes": [
                  {"at": 0, "value": 1.0}, {"at": 10, "value": 0.3}]}]}
    points = [{"at": 10, "value": 0.3}]
    assert verify_keyframes(report, "rev1", points, 1.0)
    assert not verify_keyframes(report, "rev1", points, 1.0,
                                expected_item_id="bgm1", expected_start_frame=100)
    for key in report["animations"][0]["keyframes"]:
        key["absolute_frame"] = 100 + key["at"]
    assert verify_keyframes(report, "rev1", points, 1.0,
                            expected_item_id="bgm1", expected_start_frame=100)
    report["animations"][0]["keyframes"][1]["absolute_frame"] = 111
    assert not verify_keyframes(report, "rev1", points, 1.0,
                                expected_item_id="bgm1", expected_start_frame=100)


@pytest.mark.parametrize("control", [
    "\u034f", "\u180e", "\u2061", "\u2062", "\u2063", "\u2064",
])
def test_rejects_additional_invisible_identifier_spoofing(control):
    assert not valid_identifier("item" + control + "1")
    assert not valid_identifier("rev" + control + "1", max_length=256)


def test_additional_invisible_revision_fails_readback_verification():
    report = {"success": True, "revision": "rev\u2063x", "animations": []}
    assert not verify_keyframes(report, "rev\u2063x", [], 1.0)


@pytest.mark.parametrize("control", [
    "\u00ad",  # SOFT HYPHEN: invisible in ordinary identifier rendering
    "\u0600",  # ARABIC NUMBER SIGN: format character
    "\u206a",  # INHIBIT SYMMETRIC SWAPPING
    "\u206f",  # NOMINAL DIGIT SHAPES
])
def test_rejects_other_unicode_format_controls(control):
    assert not valid_identifier("bgm" + control + "1")
    assert not valid_identifier("rev" + control + "1", max_length=256)


def test_rejects_soft_hyphen_revision_in_host_readback():
    revision = "rev\u00ad1"
    report = {"success": True, "revision": revision, "animations": []}
    assert not verify_keyframes(report, revision, [], 1.0)
