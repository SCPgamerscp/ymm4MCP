"""Read-only view of timeline items grouped by caller-defined scene ranges."""

from editing import MAX_FRAME, integer


def _uncovered(intervals: list[tuple[int, int]], start: int, end: int) -> list[dict]:
    cursor, gaps = start, []
    for left, right in sorted(intervals):
        if left > cursor:
            gaps.append({"start_frame": cursor, "end_frame": left})
        cursor = max(cursor, right)
    if cursor < end:
        gaps.append({"start_frame": cursor, "end_frame": end})
    return gaps


def _valid_identifier(value: object) -> bool:
    """Reject blank or control-character IDs without changing readable IDs."""
    return (isinstance(value, str) and bool(value.strip())
            and not any(ord(char) < 32 or ord(char) == 127 for char in value))


def index(items: list[dict], ranges: list[dict]) -> dict:
    if not isinstance(items, list):
        raise ValueError("items must be an array")
    if not isinstance(ranges, list) or not 1 <= len(ranges) <= 200:
        raise ValueError("scene_ranges must contain 1..200 ranges")
    scenes, previous_end, ids = [], 0, set()
    for position, raw in enumerate(ranges):
        if not isinstance(raw, dict) or set(raw) != {"id", "start_frame", "end_frame"}:
            raise ValueError("scene range requires id, start_frame, end_frame")
        scene_id = raw["id"]
        if not _valid_identifier(scene_id) or len(scene_id) > 64 or scene_id in ids:
            raise ValueError("scene ids must be distinct nonempty strings of at most 64 characters")
        start = integer(raw["start_frame"], f"scene_ranges[{position}].start_frame")
        end = integer(raw["end_frame"], f"scene_ranges[{position}].end_frame", 1, MAX_FRAME)
        if start >= end or start < previous_end:
            raise ValueError("scene ranges must be ordered, non-overlapping, and nonempty")
        ids.add(scene_id)
        previous_end = end
        scenes.append({"id": scene_id, "start_frame": start, "end_frame": end,
                       "items": [], "item_count": 0})

    unassigned = []
    seen_item_ids = set()
    for raw in items:
        if not isinstance(raw, dict):
            raise ValueError("item must be an object")
        frame = integer(raw.get("frame"), "item.frame")
        length = integer(raw.get("length"), "item.length", 1)
        end = integer(frame + length, "item.end_frame", 1, MAX_FRAME)
        identity = raw.get("item_id")
        if not _valid_identifier(identity):
            raise ValueError("item_id is required")
        if identity in seen_item_ids:
            raise ValueError("duplicate item_id in timeline snapshot")
        seen_item_ids.add(identity)
        layer = raw.get("layer")
        if layer is not None:
            layer = integer(layer, "item.layer")
        memberships = 0
        for scene in scenes:
            if frame < scene["end_frame"] and end > scene["start_frame"]:
                visible_start = max(frame, scene["start_frame"])
                visible_end = min(end, scene["end_frame"])
                scene["items"].append({"item_id": identity, "frame": frame, "end_frame": end,
                                       "layer": layer, "type": raw.get("type"),
                                       "visible_start_frame": visible_start,
                                       "visible_end_frame": visible_end,
                                       "relative_start_frame": visible_start - scene["start_frame"],
                                       "relative_end_frame": visible_end - scene["start_frame"],
                                       "crosses_boundary": frame < scene["start_frame"] or end > scene["end_frame"]})
                scene["item_count"] += 1
                memberships += 1
        if not memberships:
            unassigned.append(identity)
    for scene_position, scene in enumerate(scenes):
        # A continuation must actually intersect the adjacent scene. Merely
        # crossing the current scene's edge (or ending in a gap) is not enough.
        previous_scene = scenes[scene_position - 1] if scene_position else None
        next_scene = scenes[scene_position + 1] if scene_position + 1 < len(scenes) else None
        scene["continues_from_previous_item_ids"] = [
            item["item_id"] for item in scene["items"]
            if previous_scene is not None
            and item["frame"] < previous_scene["end_frame"]
            and item["end_frame"] > previous_scene["start_frame"]
        ]
        scene["continues_into_next_item_ids"] = [
            item["item_id"] for item in scene["items"]
            if next_scene is not None
            and item["frame"] < next_scene["end_frame"]
            and item["end_frame"] > next_scene["start_frame"]
        ]
        intervals = [(item["visible_start_frame"], item["visible_end_frame"])
                     for item in scene["items"]]
        gaps = _uncovered(intervals, scene["start_frame"], scene["end_frame"])
        scene["uncovered_ranges"] = gaps
        scene["uncovered_frames"] = sum(g["end_frame"] - g["start_frame"] for g in gaps)
        scene["covered_frames"] = scene["end_frame"] - scene["start_frame"] - scene["uncovered_frames"]
        by_layer = {}
        for item in scene["items"]:
            by_layer.setdefault(item["layer"], []).append(item)
        scene["layer_occupancy"] = []
        # Keep layer occupancy deterministic when the host changes item order.
        # A missing layer is sorted after all numbered layers.
        for layer in sorted(by_layer, key=lambda value: (value is None, value if value is not None else 0)):
            members = by_layer[layer]
            layer_gaps = _uncovered([(item["visible_start_frame"], item["visible_end_frame"])
                                     for item in members], scene["start_frame"], scene["end_frame"])
            uncovered = sum(g["end_frame"] - g["start_frame"] for g in layer_gaps)
            scene["layer_occupancy"].append({
                "layer": layer, "item_ids": list(dict.fromkeys(item["item_id"] for item in members)),
                "covered_frames": scene["end_frame"] - scene["start_frame"] - uncovered,
                "uncovered_frames": uncovered, "uncovered_ranges": layer_gaps,
            })
    return {"success": True, "scenes": scenes, "unassigned_item_ids": unassigned,
            "item_count": len(items), "note": "Scene ranges are caller-defined and are not saved in YMM4"}
