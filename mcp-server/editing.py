"""Side-effect-free planning and verification for timeline edits."""
import math

MAX_FRAME = 2_147_483_647


def integer(value, name, minimum=0, maximum=MAX_FRAME):
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise ValueError(f"{name} must be an integer in {minimum}..{maximum}")
    return value


def plan_script(args):
    lines = args.get("lines", [])
    if not isinstance(lines, list) or not 1 <= len(lines) <= 500:
        raise ValueError("lines must contain 1..500 dialogue entries")
    fps = integer(args.get("fps", 30), "fps", 1, 240)
    speed = args.get("chars_per_sec", 5)
    if isinstance(speed, bool) or not isinstance(speed, (int, float)) or not math.isfinite(speed) or not 0 < speed <= 1000:
        raise ValueError("chars_per_sec must be finite and in (0, 1000]")
    frame = integer(args.get("start_frame", 0), "start_frame")
    gap = integer(args.get("gap", 0), "gap")
    if "dry_run" in args and not isinstance(args["dry_run"], bool):
        raise ValueError("dry_run must be boolean")
    used_layers = set()
    for index, line in enumerate(lines):
        if not isinstance(line, dict):
            raise ValueError(f"lines[{index}] must be an object")
        if set(line) - {"text", "character", "layer"}:
            raise ValueError(f"lines[{index}] contains unsupported fields")
        for key in ("text", "character"):
            if not isinstance(line.get(key), str) or not line[key].strip():
                raise ValueError(f"lines[{index}].{key} is required")
        if len(line["text"]) > 10000:
            raise ValueError(f"lines[{index}].text exceeds 10000 characters")
        if "layer" in line:
            used_layers.add(integer(line["layer"], f"lines[{index}].layer"))
    assigned = {}
    next_layer = 0
    planned = []
    for index, line in enumerate(lines):
        layer = line.get("layer")
        if layer is None:
            if line["character"] not in assigned:
                while next_layer in used_layers:
                    next_layer += 1
                assigned[line["character"]] = next_layer
                used_layers.add(next_layer)
            layer = assigned[line["character"]]
        seconds = max(1.0, len(line["text"]) / speed)
        if not math.isfinite(seconds) or seconds > MAX_FRAME / fps:
            raise ValueError("estimated duration exceeds the frame range")
        length = max(1, math.ceil(seconds * fps))
        integer(frame + length + gap, "estimated end frame")
        planned.append({**line, "layer": layer, "frame": frame, "length": length,
                        "length_source": "estimated", "line_index": index})
        frame += length + gap
    return {"success": True, "dry_run": True, "estimated": True, "added": 0,
            "total_frames": frame, "details": planned,
            "note": "No edits or voice synthesis performed. Actual voice durations may differ."}


def _item_ids(items, indices):
    """Return only stable IDs present on the affected items."""
    return [items[index]["item_id"] for index in indices
            if isinstance(items[index], dict) and isinstance(items[index].get("item_id"), str)]


def validate_timeline(items, expected=None, duration=None, include_gaps=True):
    """Produce machine-readable structural QA without changing the timeline."""
    if not isinstance(items, list):
        raise ValueError("items must be an array")
    if duration is not None:
        integer(duration, "duration", 1)
    if not isinstance(include_gaps, bool):
        raise ValueError("include_gaps must be boolean")

    problems = []
    layers = {}
    for index, item in enumerate(items):
        try:
            if not isinstance(item, dict):
                raise ValueError("item must be an object")
            frame = integer(item.get("frame"), "frame")
            length = integer(item.get("length"), "length", 1)
            layer = integer(item.get("layer"), "layer")
            end = integer(frame + length, "end frame")
        except ValueError as exc:
            ids = _item_ids(items, [index])
            problems.append({"code": "INVALID_ITEM", "severity": "error", "index": index,
                             "item_ids": ids, "message": str(exc)})
            continue
        layers.setdefault(layer, []).append((frame, end, index))
        if duration is not None and end > duration:
            problems.append({"code": "EXCEEDS_DURATION", "severity": "error", "index": index,
                             "item_ids": _item_ids(items, [index]), "frame_range": [frame, end],
                             "project_duration": duration,
                             "suggested_fix": {"action": "edit_item", "sub_action": "property",
                                               "item_id": item.get("item_id"),
                                               "prop": "Frame" if frame >= duration else "Length",
                                               "value": max(0, duration - 1) if frame >= duration else duration - frame}})

    for layer, spans in layers.items():
        active_end = None
        active_index = None
        for frame, tail, index in sorted(spans):
            if active_index is not None and frame < active_end:
                affected = [active_index, index]
                problems.append({"code": "OVERLAP", "severity": "error", "layer": layer,
                                 "indices": affected, "item_ids": _item_ids(items, affected),
                                 "frame_range": [frame, min(active_end, tail)],
                                 "suggested_fix": {"action": "edit_item", "sub_action": "resolve_overlaps",
                                                   "layers": [layer]}})
            elif include_gaps and active_index is not None and frame > active_end:
                problems.append({"code": "GAP", "severity": "warning", "layer": layer,
                                 "indices": [active_index, index],
                                 "item_ids": _item_ids(items, [active_index, index]),
                                 "frame_range": [active_end, frame]})
            if active_end is None or tail > active_end:
                active_end, active_index = tail, index

    if expected is not None:
        if not isinstance(expected, list) or len(expected) > 1000:
            raise ValueError("expected must be an array of at most 1000 item descriptions")
        supported = {"frame", "layer", "length", "type", "text", "item_id", "revision", "id"}
        for index, wanted in enumerate(expected):
            if not isinstance(wanted, dict) or not wanted or set(wanted) - supported:
                raise ValueError("expected entries must contain supported nonempty item selectors")
            selector = dict(wanted)
            legacy_id = selector.pop("id", None)
            if legacy_id is not None:
                if "item_id" in selector and selector["item_id"] != legacy_id:
                    raise ValueError("expected id and item_id must not conflict")
                selector["item_id"] = legacy_id
            matches = [i for i, item in enumerate(items)
                       if isinstance(item, dict) and all(item.get(k) == v for k, v in selector.items())]
            if len(matches) != 1:
                problems.append({"code": "EXPECTED_NOT_FOUND" if not matches else "EXPECTED_AMBIGUOUS",
                                 "severity": "error", "expected_index": index,
                                 "expected": selector, "matches": matches,
                                 "item_ids": _item_ids(items, matches)})

    error_count = sum(problem["severity"] == "error" for problem in problems)
    warning_count = sum(problem["severity"] == "warning" for problem in problems)
    return {"success": True, "passed": error_count == 0, "valid": error_count == 0,
            "score": max(0, 100 - error_count * 20 - warning_count * 5),
            "item_count": len(items), "issue_count": len(problems),
            "summary": {"errors": error_count, "warnings": warning_count},
            "issues": problems, "problems": problems,
            "scope": "Frame ranges, same-layer overlaps/internal gaps and explicit expected items only; visual/audio quality is not checked."}
