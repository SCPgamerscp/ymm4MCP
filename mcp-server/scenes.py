"""Read-only view of timeline items grouped by caller-defined scene ranges."""

from editing import MAX_FRAME, integer


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
        if not isinstance(scene_id, str) or not scene_id or len(scene_id) > 64 or scene_id in ids:
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
    for raw in items:
        if not isinstance(raw, dict):
            raise ValueError("item must be an object")
        frame = integer(raw.get("frame"), "item.frame")
        length = integer(raw.get("length"), "item.length", 1)
        end = integer(frame + length, "item.end_frame", 1, MAX_FRAME)
        identity = raw.get("item_id")
        if not isinstance(identity, str) or not identity:
            raise ValueError("item_id is required")
        memberships = 0
        for scene in scenes:
            if frame < scene["end_frame"] and end > scene["start_frame"]:
                visible_start = max(frame, scene["start_frame"])
                visible_end = min(end, scene["end_frame"])
                scene["items"].append({"item_id": identity, "frame": frame, "end_frame": end,
                                       "layer": raw.get("layer"), "type": raw.get("type"),
                                       "visible_start_frame": visible_start,
                                       "visible_end_frame": visible_end,
                                       "relative_start_frame": visible_start - scene["start_frame"],
                                       "relative_end_frame": visible_end - scene["start_frame"],
                                       "crosses_boundary": frame < scene["start_frame"] or end > scene["end_frame"]})
                scene["item_count"] += 1
                memberships += 1
        if not memberships:
            unassigned.append(identity)
    return {"success": True, "scenes": scenes, "unassigned_item_ids": unassigned,
            "item_count": len(items), "note": "Scene ranges are caller-defined and are not saved in YMM4"}
