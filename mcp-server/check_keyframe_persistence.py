"""Read-only, two-stage YMM4 keyframe save/reopen acceptance check (#140).

Capture *after* editing Zoom/X but *before* saving. Save and reopen the
throwaway project in YMM4 manually, then verify. This tool never edits YMM4.
"""
from __future__ import annotations

import argparse
import json
import math
import os
from pathlib import Path
import sys

import httpx

from check_keyframe_compat import (ProbeError, api, keyframes,
                                   snapshot_all_animations, validate_test_item)
from ymm4_connection import connection_settings

SCHEMA = 1
MAX_SNAPSHOT_BYTES = 1_048_576


def _project(client: httpx.Client, base: str, headers: dict) -> tuple[str, Path, bool]:
    info = api(client, base, headers, "GET", "/project")
    raw = info.get("projectPath")
    if not isinstance(raw, str) or not raw.strip() or not os.path.isabs(raw):
        raise ProbeError("Project must be saved to an absolute .ymmp path")
    path = Path(raw).resolve(strict=True)
    if path.suffix.lower() != ".ymmp" or not path.is_file() or path.stat().st_size == 0:
        raise ProbeError("Project .ymmp file is missing or empty")
    saved = info.get("isSaved")
    if type(saved) is not bool:
        raise ProbeError("Project save state is unknown")
    return os.path.normcase(str(path)), path, saved


def _stable_image(client: httpx.Client, base: str, headers: dict,
                  item_id: str, at: int) -> None:
    validate_test_item(client, base, headers, item_id, at)
    # /items does not include a success flag on current YMM4 builds.
    try:
        raw = client.get(base + "/items", headers=headers)
        raw.raise_for_status()
        payload = raw.json()
    except (httpx.HTTPError, ValueError) as exc:
        raise ProbeError(f"Cannot inspect persistent item identity: {exc}") from exc
    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        raise ProbeError("Malformed timeline item inventory")
    matches = [item for item in items if isinstance(item, dict)
               and item.get("item_id") == item_id]
    if len(matches) != 1 or matches[0].get("identity_persistent") is not True:
        raise ProbeError("Test item must have a unique persistent ID across project reloads")


def _state(client: httpx.Client, base: str, headers: dict,
           item_id: str, at: int, values: dict[str, float]) -> dict:
    _stable_image(client, base, headers, item_id, at)
    animations, revision = snapshot_all_animations(client, base, headers, item_id)
    for prop, expected in values.items():
        frames, prop_revision = keyframes(client, base, headers, item_id, prop)
        if prop_revision != revision or frames != animations.get(prop):
            raise ProbeError(f"{prop}: timeline changed during snapshot")
        matches = [f for f in frames if f["at"] == at]
        if len(matches) != 1 or type(matches[0].get("value")) not in (int, float):
            raise ProbeError(f"{prop}: expected keyframe is missing or ambiguous")
        actual = matches[0]["value"]
        if not math.isfinite(actual) or not math.isclose(actual, expected, rel_tol=1e-5, abs_tol=1e-5):
            raise ProbeError(f"{prop}: expected {expected}, got {actual!r}")
    return animations


def capture(client: httpx.Client, base: str, headers: dict, item_id: str,
            at: int, values: dict[str, float], snapshot: Path) -> None:
    project_name, path, saved = _project(client, base, headers)
    if saved:
        raise ProbeError("Edit the disposable project before capture; isSaved must be false")
    animations = _state(client, base, headers, item_id, at, values)
    # Guard against a concurrent project switch or save while reading.
    second_name, _, second_saved = _project(client, base, headers)
    if second_name != project_name or second_saved:
        raise ProbeError("Project changed or was saved during capture")
    payload = {"schema": SCHEMA, "project_path": project_name,
               "project_mtime_ns": path.stat().st_mtime_ns,
               "item_id": item_id, "at": at, "values": values,
               "animations": animations}
    encoded = (json.dumps(payload, ensure_ascii=False, allow_nan=False,
                          separators=(",", ":")) + "\n").encode("utf-8")
    if len(encoded) > MAX_SNAPSHOT_BYTES:
        raise ProbeError("Animation snapshot is too large")
    try:
        fd = os.open(snapshot, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "wb") as output:
            output.write(encoded)
    except FileExistsError as exc:
        raise ProbeError("Snapshot already exists; refusing overwrite") from exc
    print("PASS capture: unsaved Zoom/X keys recorded; now save and reopen the disposable project")


def verify(client: httpx.Client, base: str, headers: dict, snapshot: Path) -> None:
    if snapshot.stat().st_size > MAX_SNAPSHOT_BYTES:
        raise ProbeError("Snapshot file exceeds size limit")
    try:
        data = json.loads(snapshot.read_text(encoding="utf-8"))
    except (ValueError, UnicodeError) as exc:
        raise ProbeError("Snapshot is not valid UTF-8 JSON") from exc
    if (not isinstance(data, dict) or data.get("schema") != SCHEMA
            or not isinstance(data.get("project_path"), str)
            or type(data.get("project_mtime_ns")) is not int
            or not isinstance(data.get("item_id"), str)
            or type(data.get("at")) is not int or data["at"] <= 0
            or not isinstance(data.get("values"), dict)
            or set(data["values"]) != {"Zoom", "X"}
            or any(type(v) not in (int, float) or not math.isfinite(v)
                   for v in data["values"].values())
            or not isinstance(data.get("animations"), dict)):
        raise ProbeError("Snapshot has an invalid schema")
    project_name, path, saved = _project(client, base, headers)
    if project_name != data["project_path"] or not saved:
        raise ProbeError("Project path changed or project still has unsaved edits")
    if path.stat().st_mtime_ns <= data["project_mtime_ns"]:
        raise ProbeError("Project file was not saved after capture")
    current = _state(client, base, headers, data["item_id"],
                     data["at"], data["values"])
    if current != data["animations"]:
        raise ProbeError("Animation inventory differs after save/reopen")
    print("PASS verify: saved project and full Zoom/X animation inventory match capture")
    print("NOTE: YMM4 reopen was manually confirmed, not observable through this API")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("capture", "verify"))
    parser.add_argument("--snapshot", type=Path, required=True)
    parser.add_argument("--item-id", help="persistent image item ID (capture only)")
    parser.add_argument("--at", type=int, help="relative keyframe (capture only)")
    parser.add_argument("--zoom", type=float, help="expected Zoom value (capture only)")
    parser.add_argument("--x", type=float, help="expected X value (capture only)")
    parser.add_argument("--confirm-test-project", action="store_true")
    parser.add_argument("--confirm-reopened", action="store_true")
    args = parser.parse_args(argv)
    if not args.confirm_test_project:
        parser.error("requires --confirm-test-project (disposable project only)")
    if args.action == "capture" and (not args.item_id or args.at is None or args.at <= 0
            or args.zoom is None or args.x is None
            or not math.isfinite(args.zoom) or not math.isfinite(args.x)):
        parser.error("capture requires --item-id, --at>0, --zoom and --x (finite)")
    if args.action == "verify" and not args.confirm_reopened:
        parser.error("verify requires --confirm-reopened after manually saving and reopening")
    try:
        base, headers = connection_settings()
        with httpx.Client(timeout=10.0, trust_env=False) as client:
            if args.action == "capture":
                capture(client, base, headers, args.item_id, args.at,
                        {"Zoom": args.zoom, "X": args.x}, args.snapshot)
            else:
                verify(client, base, headers, args.snapshot)
    except (ProbeError, OSError, ValueError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
