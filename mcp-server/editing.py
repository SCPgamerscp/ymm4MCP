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


def validate_timeline(items, expected=None, duration=None):
    if not isinstance(items, list):
        raise ValueError("items must be an array")
    if duration is not None:
        integer(duration, "duration", 1)
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
            problems.append({"code": "INVALID_ITEM", "severity": "error", "index": index, "message": str(exc)})
            continue
        layers.setdefault(layer, []).append((frame, end, index))
        if duration is not None and end > duration:
            problems.append({"code": "EXCEEDS_DURATION", "severity": "error", "index": index, "end_frame": end})
    for layer, spans in layers.items():
        end, previous = 0, None
        for frame, tail, index in sorted(spans):
            if previous is not None and frame < end:
                problems.append({"code": "OVERLAP", "severity": "error", "layer": layer,
                                 "indices": [previous, index], "start_frame": frame, "end_frame": min(end, tail)})
            elif frame > end:
                problems.append({"code": "GAP", "severity": "warning", "layer": layer,
                                 "start_frame": end, "end_frame": frame})
            if tail > end:
                end, previous = tail, index
    if expected is not None:
        if not isinstance(expected, list) or len(expected) > 1000:
            raise ValueError("expected must be an array of at most 1000 item descriptions")
        for index, wanted in enumerate(expected):
            if not isinstance(wanted, dict) or not wanted or set(wanted) - {"frame", "layer", "length", "type", "text", "id"}:
                raise ValueError("expected entries must contain supported nonempty item selectors")
            matches = [i for i, item in enumerate(items) if isinstance(item, dict) and all(item.get(k) == v for k, v in wanted.items())]
            if len(matches) != 1:
                problems.append({"code": "EXPECTED_NOT_FOUND" if not matches else "EXPECTED_AMBIGUOUS",
                                 "severity": "error", "expected_index": index, "matches": matches})
    return {"success": True, "valid": not any(p["severity"] == "error" for p in problems),
            "item_count": len(items), "problems": problems,
            "scope": "Frame ranges, same-layer overlaps/gaps and explicit expected items only; visual/audio quality is not checked."}
