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
import stat
import sys

import httpx

from check_keyframe_compat import (ProbeError, api, keyframes,
                                   snapshot_all_animations, valid_timeline_id,
                                   validate_test_item)
from ymm4_connection import connection_settings

SCHEMA = 1
MAX_SNAPSHOT_BYTES = 1_048_576


def _project(client: httpx.Client, base: str, headers: dict) -> tuple[str, os.stat_result, bool]:
    info = api(client, base, headers, "GET", "/project")
    raw = info.get("projectPath")
    if not isinstance(raw, str) or not raw.strip() or not os.path.isabs(raw):
        raise ProbeError("Project must be saved to an absolute .ymmp path")
    path = Path(raw).resolve(strict=True)
    file_info = path.stat()
    if path.suffix.lower() != ".ymmp" or not stat.S_ISREG(file_info.st_mode) or file_info.st_size == 0:
        raise ProbeError("Project .ymmp file is missing or empty")
    saved = info.get("isSaved")
    if type(saved) is not bool:
        raise ProbeError("Project save state is unknown")
    return os.path.normcase(str(path)), file_info, saved


def _same_project_file(before: os.stat_result, after: os.stat_result) -> bool:
    # File replacement can preserve its name, size, and modification time.
    return (os.path.samestat(before, after)
            and (before.st_size, before.st_mtime_ns, before.st_ctime_ns)
            == (after.st_size, after.st_mtime_ns, after.st_ctime_ns))


def _state(client: httpx.Client, base: str, headers: dict,
           item_id: str, at: int | None, values: dict[str, float]) -> tuple[dict, int, str | None]:
    initial = validate_test_item(client, base, headers, item_id, at, require_persistent=True)
    resolved_at, item_revision, timeline_id = initial
    animations, revision = snapshot_all_animations(client, base, headers, item_id)
    if item_revision != revision:
        raise ProbeError("Image revision changed or was unavailable during snapshot")
    for prop, expected in values.items():
        frames, prop_revision = keyframes(client, base, headers, item_id, prop)
        if prop_revision != revision or frames != animations.get(prop):
            raise ProbeError(f"{prop}: timeline changed during snapshot")
        matches = [f for f in frames if f["at"] == resolved_at]
        if len(matches) != 1 or type(matches[0].get("value")) not in (int, float):
            raise ProbeError(f"{prop}: expected keyframe is missing or ambiguous")
        actual = matches[0]["value"]
        if not math.isfinite(actual) or not math.isclose(actual, expected, rel_tol=1e-5, abs_tol=1e-5):
            raise ProbeError(f"{prop}: expected {expected}, got {actual!r}")
    if validate_test_item(client, base, headers, item_id, at, require_persistent=True) != initial:
        raise ProbeError("Image or live timeline changed during snapshot")
    return animations, resolved_at, timeline_id


def capture(client: httpx.Client, base: str, headers: dict, item_id: str,
            at: int | None, values: dict[str, float], snapshot: Path) -> None:
    project_name, project_file, saved = _project(client, base, headers)
    if saved:
        raise ProbeError("Edit the disposable project before capture; isSaved must be false")
    last_frame = at is None
    animations, at, timeline_id = _state(client, base, headers, item_id, at, values)
    # Guard against a concurrent project switch or save while reading.
    second_name, second_file, second_saved = _project(client, base, headers)
    if second_name != project_name or second_saved or not _same_project_file(project_file, second_file):
        raise ProbeError("Project changed or was saved during capture")
    payload = {"schema": SCHEMA, "project_path": project_name,
               "project_mtime_ns": project_file.st_mtime_ns,
               "item_id": item_id, "at": at, "last_frame": last_frame, "values": values,
               "animations": animations}
    if timeline_id is not None:
        payload["timeline_id"] = timeline_id
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
    print(f"PASS capture: unsaved Zoom/X keys recorded at frame {at}; now save and reopen the disposable project")


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
            or type(data.get("last_frame", False)) is not bool
            or ("timeline_id" in data and not valid_timeline_id(data["timeline_id"]))
            or not isinstance(data.get("values"), dict)
            or set(data["values"]) != {"Zoom", "X"}
            or any(type(v) not in (int, float) or not math.isfinite(v)
                   for v in data["values"].values())
            or not isinstance(data.get("animations"), dict)):
        raise ProbeError("Snapshot has an invalid schema")
    project_name, project_file, saved = _project(client, base, headers)
    if project_name != data["project_path"] or not saved:
        raise ProbeError("Project path changed or project still has unsaved edits")
    if project_file.st_mtime_ns <= data["project_mtime_ns"]:
        raise ProbeError("Project file was not saved after capture")
    current, at, timeline_id = _state(client, base, headers, data["item_id"],
                                      None if data.get("last_frame") else data["at"],
                                      data["values"])
    previous_timeline_id = data.get("timeline_id")
    if previous_timeline_id is not None:
        if timeline_id is None:
            raise ProbeError("Live timeline identity was present at capture but is now missing")
        if timeline_id == previous_timeline_id:
            raise ProbeError("The original live timeline is still active; close and reopen the saved project")
    if at != data["at"]:
        raise ProbeError("Image last frame differs after save/reopen")
    if current != data["animations"]:
        raise ProbeError("Animation inventory differs after save/reopen")
    second_name, second_file, second_saved = _project(client, base, headers)
    if second_name != project_name or not second_saved or not _same_project_file(project_file, second_file):
        raise ProbeError("Project changed during save/reopen verification")
    print("PASS verify: saved project and full Zoom/X animation inventory match capture")
    if previous_timeline_id is not None:
        print("PASS verify: live timeline identity changed after the operator-confirmed reopen")
    else:
        print("NOTE: Capture has no live timeline identity; YMM4 reopen relies on manual confirmation")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("capture", "verify"))
    parser.add_argument("--snapshot", type=Path, required=True)
    parser.add_argument("--item-id", help="persistent image item ID (capture only)")
    frames = parser.add_mutually_exclusive_group()
    frames.add_argument("--at", type=int, help="relative keyframe (capture only)")
    frames.add_argument("--last-frame", action="store_true", help="use the target image's last frame (capture only)")
    parser.add_argument("--zoom", type=float, help="expected Zoom value (capture only)")
    parser.add_argument("--x", type=float, help="expected X value (capture only)")
    parser.add_argument("--confirm-test-project", action="store_true")
    parser.add_argument("--confirm-reopened", action="store_true")
    args = parser.parse_args(argv)
    if not args.confirm_test_project:
        parser.error("requires --confirm-test-project (disposable project only)")
    if args.action == "capture" and (not args.item_id
            or (not args.last_frame and (args.at is None or args.at <= 0))
            or args.zoom is None or args.x is None
            or not math.isfinite(args.zoom) or not math.isfinite(args.x)):
        parser.error("capture requires --item-id, --at>0 or --last-frame, and finite --zoom and --x")
    if args.action == "verify" and args.last_frame:
        parser.error("--last-frame is only used at capture")
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
