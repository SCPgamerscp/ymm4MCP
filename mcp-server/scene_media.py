"""Estimate independent visual and audio coverage from a read-only scene index."""

import re

from editing import MAX_FRAME

MAX_SCENE_ITEMS = 50_000
MAX_SCENES = 200


def is_known_item_type(value: object, names: tuple[str, ...]) -> bool:
    """Match the exact CLR type name, optionally with namespace/nesting qualifiers."""
    # An empty or malformed namespace component must not spoof a known
    # rendering type (for example, "Other..VideoItem").
    if (not isinstance(value, str) or len(value) > 128 or
            re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*(?:[.+][A-Za-z_][A-Za-z0-9_]*)*",
                         value) is None):
        return False
    leaf = value.rsplit(".", 1)[-1].rsplit("+", 1)[-1]
    return leaf.lower() in names


def _scene_range(scene: dict) -> tuple[int, int]:
    start, end = scene["start_frame"], scene["end_frame"]
    if (type(start) is not int or type(end) is not int or
            start < 0 or end > MAX_FRAME or end <= start):
        raise ValueError("scene range must have bounded integer frames and positive duration")
    return start, end


def _spans(scene: dict, names: tuple[str, ...], start: int, end: int) -> list[tuple[int, int]]:
    spans = []
    for item in scene["items"]:
        if not isinstance(item, dict):
            raise ValueError("scene item must be an object")
        if not is_known_item_type(item.get("type"), names):
            continue
        left, right = item["visible_start_frame"], item["visible_end_frame"]
        if type(left) is not int or type(right) is not int or right <= left:
            raise ValueError("visible item range must have integer frames and positive duration")
        # Items may extend beyond a scene, but cannot cover frames outside it.
        left, right = max(start, left), min(end, right)
        if left < right:
            spans.append((left, right))
    return sorted(spans)


def with_coverage(index: dict) -> dict:
    """Add coverage only after validating the complete scene evidence."""
    if (not isinstance(index, dict) or not isinstance(index.get("scenes"), list)
            or not 1 <= len(index["scenes"]) <= MAX_SCENES):
        raise ValueError("scene index requires 1..200 scenes")
    seen_ids = set()
    for scene in index["scenes"]:
        if not isinstance(scene, dict):
            raise ValueError("scene must be an object")
        items = scene.get("items")
        if not isinstance(items, list) or len(items) > MAX_SCENE_ITEMS:
            raise ValueError("scene items must be a bounded array")
        item_ids = set()
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("scene item must be an object")
            if "item_id" in item:
                identity = item["item_id"]
                if (not isinstance(identity, str) or not identity.strip() or
                        len(identity) > 256 or identity in item_ids):
                    raise ValueError("scene item IDs must be distinct nonempty strings")
                item_ids.add(identity)
            if "frame" in item or "end_frame" in item:
                frame, source_end = item.get("frame"), item.get("end_frame")
                if (type(frame) is not int or type(source_end) is not int or
                        frame < 0 or source_end <= frame or source_end > MAX_FRAME):
                    raise ValueError("scene source item range is invalid")
                visible_start = item.get("visible_start_frame")
                visible_end = item.get("visible_end_frame")
                start, end = _scene_range(scene)
                if (visible_start != max(start, frame) or
                        visible_end != min(end, source_end) or
                        type(visible_start) is not int or type(visible_end) is not int):
                    raise ValueError("scene visible range contradicts source item")
        if "item_count" in scene and (type(scene["item_count"]) is not int or
                                       scene["item_count"] != len(items)):
            raise ValueError("scene item_count does not match item evidence")
        if "id" in scene:
            identifier = scene["id"]
            if (not isinstance(identifier, str) or not identifier.strip() or
                    len(identifier) > 64 or identifier in seen_ids):
                raise ValueError("scene IDs must be distinct nonempty strings")
            seen_ids.add(identifier)
    # Unknown host item types are excluded rather than assumed to provide a picture.
    categories = {"visual": ("videoitem", "imageitem", "textitem", "tachieitem",
                             "tachiefaceitem", "faceitem"),
                  "audio": ("audioitem", "voiceitem")}
    pending = []
    for scene in index["scenes"]:
        start, end = _scene_range(scene)
        updates = {}
        for category, names in categories.items():
            cursor = start
            gaps = []
            for left, right in _spans(scene, names, start, end):
                if left > cursor:
                    gaps.append({"start_frame": cursor, "end_frame": left})
                cursor = max(cursor, right)
            if cursor < end:
                gaps.append({"start_frame": cursor, "end_frame": end})
            duration = end - start
            covered = duration - sum(g["end_frame"] - g["start_frame"] for g in gaps)
            updates[f"{category}_uncovered_ranges"] = gaps
            updates[f"{category}_covered_frames"] = covered
            updates[f"{category}_coverage_ratio"] = covered / duration
        pending.append((scene, updates))
    # Do not leave a partly updated index when a later scene contains invalid data.
    for scene, updates in pending:
        scene.update(updates)
    return index
