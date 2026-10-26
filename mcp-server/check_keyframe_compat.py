"""Opt-in YMM4 4.56.x keyframe end-to-end smoke test for Issue #140.

Run only on a disposable copy of a project with a test image item.
The script adds and then removes *new* keyframes; it does not save the project.
It deliberately refuses to overwrite existing keyframes at the test frame.
"""
from __future__ import annotations

import argparse
import math
import sys
from urllib.parse import quote

import httpx

from ymm4_connection import connection_settings


class ProbeError(RuntimeError):
    pass


def api(client: httpx.Client, base: str, headers: dict[str, str],
        method: str, path: str, payload: dict | None = None) -> dict:
    try:
        response = client.request(method, base + path, headers=headers, json=payload)
        response.raise_for_status()
        result = response.json()
    except (httpx.HTTPError, ValueError) as exc:
        raise ProbeError(f"{method} {path}: HTTP/JSON failure: {exc}") from exc
    if not isinstance(result, dict) or result.get("success") is not True:
        code = result.get("error_code") if isinstance(result, dict) else "INVALID_RESPONSE"
        raise ProbeError(f"{method} {path}: plugin rejected request ({code})")
    return result


def keyframes(client: httpx.Client, base: str, headers: dict[str, str],
              item_id: str, prop: str) -> tuple[list[dict], str]:
    path = f"/items/keyframes?item_id={quote(item_id, safe='')}&prop={quote(prop, safe='')}"
    result = api(client, base, headers, "GET", path)
    animations = result.get("animations")
    if not isinstance(animations, list):
        raise ProbeError(f"{prop}: malformed animation response")
    matches = [a for a in animations
               if isinstance(a, dict) and a.get("prop") == prop]
    if len(matches) != 1 or not isinstance(matches[0].get("keyframes"), list):
        raise ProbeError(f"{prop}: Animation or keyframes unavailable")
    frames = matches[0]["keyframes"]
    if any(not isinstance(f, dict) or type(f.get("at")) is not int
           or f["at"] < 0 for f in frames):
        raise ProbeError(f"{prop}: malformed keyframe readback")
    revision = result.get("revision")
    if not isinstance(revision, str) or not revision:
        raise ProbeError(f"{prop}: revision unavailable; refusing unsafe edit")
    return frames, revision


def has_frame(frames: list[dict], at: int) -> bool:
    return any(frame["at"] == at for frame in frames)


def run_probe(client: httpx.Client, base: str, headers: dict[str, str],
              item_id: str, at: int, values: dict[str, float]) -> None:
    # Validate every property before the first write, so a missing X/Zoom
    # cannot leave the other property modified.
    baseline: dict[str, list[dict]] = {}
    for prop in values:
        frames, _ = keyframes(client, base, headers, item_id, prop)
        if has_frame(frames, at):
            raise ProbeError(f"{prop}: frame {at} already exists; refusing overwrite")
        # Capture all original keys, not just the proposed insertion position.
        # Shared YMM4 KeyFrames can affect other Animation properties as well.
        baseline[prop] = frames

    # KeyFrames can be shared by all animations on a YMM4 timeline item.
    # Test one property at a time: remove Zoom's frame before inserting X's,
    # otherwise the first edit can create a shared frame for the second.
    for prop, value in values.items():
        touched = False
        failure: Exception | None = None
        cleanup_error: Exception | None = None
        try:
            _, revision = keyframes(client, base, headers, item_id, prop)
            # Record BEFORE the POST: a failed request might have mutated YMM4.
            touched = True
            api(client, base, headers, "POST", "/items/keyframe",
                {"item_id": item_id, "expected_revision": revision,
                 "prop": prop, "action": "set", "at": at, "value": value})
            frames, _ = keyframes(client, base, headers, item_id, prop)
            matches = [f for f in frames if f["at"] == at]
            if len(matches) != 1:
                raise ProbeError(f"{prop}: inserted frame not uniquely visible")
            actual = matches[0].get("value")
            if (isinstance(actual, bool) or not isinstance(actual, (int, float))
                    or not math.isfinite(actual)
                    or not math.isclose(actual, value, rel_tol=1e-5, abs_tol=1e-5)):
                raise ProbeError(f"{prop}: wrong value after insertion: {actual!r}")
            print(f"PASS {prop}: inserted frame={at} value={actual}")
        except Exception as exc:
            failure = exc
        finally:
            if touched:
                try:
                    frames, revision = keyframes(client, base, headers, item_id, prop)
                    if has_frame(frames, at):
                        api(client, base, headers, "POST", "/items/keyframe",
                            {"item_id": item_id, "expected_revision": revision,
                             "prop": prop, "action": "remove", "at": at})
                    # Removing a shared frame may silently alter unrelated
                    # Animation values or pre-existing frames. A missing test
                    # frame alone is NOT proof that cleanup restored the project.
                    for checked_prop, original_frames in baseline.items():
                        restored, _ = keyframes(client, base, headers, item_id, checked_prop)
                        if restored != original_frames:
                            raise ProbeError(
                                f"{checked_prop}: original keyframes changed after "
                                f"{prop} cleanup; restore the test project from backup")
                    print(f"PASS {prop}: all original Zoom/X keyframes restored")
                except Exception as exc:
                    cleanup_error = exc
        if cleanup_error is not None:
            raise ProbeError(f"CLEANUP FAILED: {prop}: {cleanup_error}; "
                             "restore your disposable project from backup") from failure
        if failure is not None:
            raise ProbeError(str(failure)) from failure


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--item-id", required=True, help="ID of an image in a disposable project")
    parser.add_argument("--at", type=int, required=True, help="unused relative frame, > 0 and < item length")
    parser.add_argument("--zoom", type=float, required=True, help="safe test Zoom value for this YMM4 build")
    parser.add_argument("--x", type=float, required=True, help="safe test X position for this YMM4 build")
    parser.add_argument("--confirm-test-project", action="store_true",
                        help="I have backed up and opened a disposable test project")
    args = parser.parse_args(argv)
    if (not args.confirm_test_project or not args.item_id.strip()
            or not 0 < args.at <= 2_147_483_647
            or not math.isfinite(args.zoom) or not math.isfinite(args.x)):
        parser.error("requires --confirm-test-project, valid item-id, at>0, and finite values")
    try:
        base, headers = connection_settings()
        # Never send the local auth token to an HTTP(S) proxy.
        with httpx.Client(timeout=10.0, trust_env=False) as client:
            run_probe(client, base, headers, args.item_id, args.at,
                      {"Zoom": args.zoom, "X": args.x})
    except (ProbeError, ValueError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    print("PASS: Zoom/X insertion, readback and removal; save/reopen still needs manual verification")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
