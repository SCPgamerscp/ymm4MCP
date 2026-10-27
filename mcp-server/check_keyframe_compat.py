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


def snapshot_all_animations(client: httpx.Client, base: str,
                            headers: dict[str, str], item_id: str) -> tuple[dict[str, list[dict]], str]:
    """Snapshot every animation, not only Zoom/X, before any write.

    YMM4 may share one KeyFrames collection across multiple properties.
    Removing a test key can alter unrelated animations; a successful cleanup
    must restore the complete observable item state.
    """
    path = f"/items/keyframes?item_id={quote(item_id, safe='')}&strict_inventory=1"
    result = api(client, base, headers, "GET", path)
    if result.get("inventory_complete") is not True:
        raise ProbeError("Incomplete animation inventory; refusing edit")
    animations = result.get("animations")
    if not isinstance(animations, list) or not animations:
        raise ProbeError("Cannot snapshot all item animations; refusing edit")
    snapshot: dict[str, list[dict]] = {}
    for animation in animations:
        if not isinstance(animation, dict):
            raise ProbeError("Malformed animation inventory; refusing edit")
        prop, frames = animation.get("prop"), animation.get("keyframes")
        if (not isinstance(prop, str) or not prop or prop in snapshot
                or not isinstance(frames, list)
                or any(not isinstance(frame, dict)
                       or type(frame.get("at")) is not int or frame["at"] < 0
                       for frame in frames)):
            raise ProbeError("Malformed animation inventory; refusing edit")
        snapshot[prop] = frames
    revision = result.get("revision")
    if not isinstance(revision, str) or not revision:
        raise ProbeError("Animation inventory revision unavailable; refusing edit")
    return snapshot, revision


def validate_test_item(client: httpx.Client, base: str,
                       headers: dict[str, str], item_id: str, at: int) -> None:
    """Fail before writing if the selected image or relative frame is unsafe."""
    try:
        response = client.get(base + "/items", headers=headers)
        response.raise_for_status()
        payload = response.json()
    except (httpx.HTTPError, ValueError) as exc:
        raise ProbeError(f"Cannot inspect target timeline item: {exc}") from exc
    if not isinstance(payload, dict) or not isinstance(payload.get("items"), list):
        raise ProbeError("Malformed timeline item inventory; refusing edit")
    matches = [item for item in payload["items"]
               if isinstance(item, dict) and item.get("item_id") == item_id]
    if len(matches) != 1:
        raise ProbeError("Target item must exist exactly once; refusing edit")
    item = matches[0]
    if "image" not in str(item.get("type", "")).lower():
        raise ProbeError("Target item is not an image; refusing edit")
    length = item.get("length")
    if type(length) is not int or length <= at:
        raise ProbeError(f"Frame {at} is outside the target item length; refusing edit")


def run_probe(client: httpx.Client, base: str, headers: dict[str, str],
              item_id: str, at: int, values: dict[str, float]) -> None:
    validate_test_item(client, base, headers, item_id, at)
    # Capture every property before writing: YMM4 may share one KeyFrames
    # collection among Zoom, X and other animation properties.
    baseline, baseline_revision = snapshot_all_animations(client, base, headers, item_id)
    for prop in values:
        if prop not in baseline:
            raise ProbeError(f"{prop}: Animation or keyframes unavailable")
        frames, revision = keyframes(client, base, headers, item_id, prop)
        if revision != baseline_revision:
            raise ProbeError(f"{prop}: item changed during preflight; refusing edit")
        if frames != baseline[prop]:
            raise ProbeError(f"{prop}: inconsistent preflight readback; refusing edit")
        if has_frame(frames, at):
            raise ProbeError(f"{prop}: frame {at} already exists; refusing overwrite")

    # KeyFrames can be shared by all animations on a YMM4 timeline item.
    # Test one property at a time: remove Zoom's frame before inserting X's,
    # otherwise the first edit can create a shared frame for the second.
    for prop, value in values.items():
        touched = False
        cleanup_revision: str | None = None
        failure: Exception | None = None
        cleanup_error: Exception | None = None
        try:
            # Check the complete inventory again before EACH write. A prior
            # probe/cleanup or another editor may have changed an unrelated
            # animation since the initial preflight. A matching Zoom/X list
            # alone is not sufficient evidence that the item is unchanged.
            current, current_revision = snapshot_all_animations(
                client, base, headers, item_id)
            if current != baseline:
                raise ProbeError(f"{prop}: animation inventory changed; refusing edit")
            _, revision = keyframes(client, base, headers, item_id, prop)
            if revision != current_revision:
                raise ProbeError(f"{prop}: item changed before write; refusing edit")
            # Record BEFORE the POST: a failed request might have mutated YMM4.
            touched = True
            result = api(client, base, headers, "POST", "/items/keyframe",
                         {"item_id": item_id, "expected_revision": revision,
                          "prop": prop, "action": "set", "at": at, "value": value})
            # Prefer the revision returned by the mutation itself: another
            # caller may edit the item before the next readback.
            returned_revision = result.get("revision")
            if isinstance(returned_revision, str) and returned_revision:
                cleanup_revision = returned_revision
            frames, readback_revision = keyframes(client, base, headers, item_id, prop)
            if cleanup_revision is not None and readback_revision != cleanup_revision:
                raise ProbeError(f"{prop}: item changed after write; refusing cleanup")
            cleanup_revision = readback_revision
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
            # The plugin's revision guard rejects the write before mutation.
            # Removing a pre-existing frame after that rejection is unsafe.
            if "REVISION_CONFLICT" in str(exc):
                touched = False
        finally:
            if touched:
                try:
                    frames, revision = keyframes(client, base, headers, item_id, prop)
                    if has_frame(frames, at):
                        if cleanup_revision is not None and revision != cleanup_revision:
                            raise ProbeError(
                                f"{prop}: item changed after readback; refusing cleanup")
                        api(client, base, headers, "POST", "/items/keyframe",
                            {"item_id": item_id,
                             "expected_revision": cleanup_revision or revision,
                             "prop": prop, "action": "remove", "at": at})
                    # Removing a shared frame may silently alter unrelated
                    # Animation values or pre-existing frames. Verify the
                    # entire inventory, including added/removed properties.
                    restored, _ = snapshot_all_animations(client, base, headers, item_id)
                    if restored != baseline:
                        changed = sorted(set(restored) ^ set(baseline))
                        changed += sorted(name for name in restored.keys() & baseline.keys()
                                          if restored[name] != baseline[name])
                        raise ProbeError(
                            f"{', '.join(changed)}: original keyframes changed after "
                            f"{prop} cleanup; restore the test project from backup")
                    print(f"PASS {prop}: all original item animations restored")
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
