"""Estimate independent visual and audio coverage from a read-only scene index."""

import re

from editing import MAX_FRAME

MAX_SCENE_ITEMS = 50_000
MAX_TOTAL_SCENE_ITEMS = 100_000
MAX_SCENES = 200


def _safe_identifier(value, limit):
    return (isinstance(value, str) and bool(value.strip()) and len(value) <= limit
            and not any((code := ord(char)) < 32 or 127 <= code <= 159
                        or code in (0x061C, 0x200E, 0x200F)
                        or 0x202A <= code <= 0x202E
                        or 0x2066 <= code <= 0x2069
                        for char in value))


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
    start, end = scene.get("start_frame"), scene.get("end_frame")
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
        left, right = item.get("visible_start_frame"), item.get("visible_end_frame")
        if type(left) is not int or type(right) is not int or right <= left:
            raise ValueError("visible item range must have integer frames and positive duration")
        # Items may extend beyond a scene, but cannot cover frames outside it.
        left, right = max(start, left), min(end, right)
        if left < right:
            spans.append((left, right))
    return sorted(spans)


def _uncovered_from_items(items: list[dict], start: int, end: int) -> list[dict]:
    """Recompute occupancy from the source spans, without trusting summaries."""
    intervals = []
    for item in items:
        left, right = item.get("visible_start_frame"), item.get("visible_end_frame")
        if type(left) is not int or type(right) is not int or right <= left:
            raise ValueError("scene occupancy requires valid visible item ranges")
        left, right = max(start, left), min(end, right)
        if left < right:
            intervals.append((left, right))
    cursor, gaps = start, []
    for left, right in sorted(intervals):
        if left > cursor:
            gaps.append({"start_frame": cursor, "end_frame": left})
        cursor = max(cursor, right)
    if cursor < end:
        gaps.append({"start_frame": cursor, "end_frame": end})
    return gaps


def _same_ranges(actual: object, expected: list[dict]) -> bool:
    """Compare frame ranges without treating JSON booleans as integers."""
    return (isinstance(actual, list) and len(actual) == len(expected)
            and all(isinstance(row, dict)
                    and set(row) == {"start_frame", "end_frame"}
                    and type(row["start_frame"]) is int
                    and type(row["end_frame"]) is int
                    and row == reference
                    for row, reference in zip(actual, expected)))


def _same_layer_occupancy(actual: object, expected: list[dict]) -> bool:
    """Check nested numeric evidence with exact types, not bool == 1."""
    if not isinstance(actual, list) or len(actual) != len(expected):
        return False
    for row, reference in zip(actual, expected):
        if not isinstance(row, dict) or set(row) != set(reference):
            return False
        if ((row["layer"] is not None and type(row["layer"]) is not int)
                or row["layer"] != reference["layer"]
                or not isinstance(row["item_ids"], list)
                or row["item_ids"] != reference["item_ids"]
                or type(row["covered_frames"]) is not int
                or type(row["uncovered_frames"]) is not int
                or row["covered_frames"] != reference["covered_frames"]
                or row["uncovered_frames"] != reference["uncovered_frames"]
                or not _same_ranges(row["uncovered_ranges"],
                                    reference["uncovered_ranges"])):
            return False
    return True


def _validate_scene_summaries(scene: dict, start: int, end: int) -> None:
    """Reject stale or forged scene-wide and per-layer occupancy summaries."""
    keys = ("covered_frames", "uncovered_frames", "uncovered_ranges", "layer_occupancy")
    if not any(key in scene for key in keys):
        return
    gaps = _uncovered_from_items(scene["items"], start, end)
    uncovered = sum(g["end_frame"] - g["start_frame"] for g in gaps)
    covered = end - start - uncovered
    expected = {"covered_frames": covered, "uncovered_frames": uncovered,
                "uncovered_ranges": gaps}
    for key, value in expected.items():
        if key not in scene:
            continue
        valid = (_same_ranges(scene[key], value) if key == "uncovered_ranges"
                 else type(scene[key]) is int and scene[key] == value)
        if not valid:
            raise ValueError(f"scene {key} contradicts item occupancy")
    if "layer_occupancy" in scene:
        groups = {}
        for item in scene["items"]:
            groups.setdefault(item.get("layer"), []).append(item)
        expected_layers = []
        for layer in sorted(groups, key=lambda x: (x is None, x if x is not None else 0)):
            members = groups[layer]
            layer_gaps = _uncovered_from_items(members, start, end)
            missing = sum(g["end_frame"] - g["start_frame"] for g in layer_gaps)
            expected_layers.append({
                "layer": layer,
                "item_ids": list(dict.fromkeys(item.get("item_id") for item in members)),
                "covered_frames": end - start - missing,
                "uncovered_frames": missing,
                "uncovered_ranges": layer_gaps,
            })
        if not _same_layer_occupancy(scene["layer_occupancy"], expected_layers):
            raise ValueError("scene layer_occupancy contradicts item occupancy")


def _validate_continuations(scenes: list[dict]) -> None:
    """Continuation IDs must match the actual overlap with adjacent scenes."""
    for pos, scene in enumerate(scenes):
        for field, neighbor in (
            ("continues_from_previous_item_ids", scenes[pos - 1] if pos else None),
            ("continues_into_next_item_ids", scenes[pos + 1] if pos + 1 < len(scenes) else None),
        ):
            if field not in scene:
                continue
            expected = []
            if neighbor is not None:
                for item in scene["items"]:
                    if (not isinstance(item.get("item_id"), str) or
                            type(item.get("frame")) is not int or
                            type(item.get("end_frame")) is not int):
                        raise ValueError("scene continuation requires complete item evidence")
                    if (item["frame"] < neighbor["end_frame"] and
                            item["end_frame"] > neighbor["start_frame"]):
                        expected.append(item["item_id"])
            if scene[field] != expected:
                raise ValueError(f"scene {field} contradicts adjacent scene membership")


def with_coverage(index: dict) -> dict:
    """Add coverage only after validating the complete scene evidence."""
    if (not isinstance(index, dict) or not isinstance(index.get("scenes"), list)
            or not 1 <= len(index["scenes"]) <= MAX_SCENES):
        raise ValueError("scene index requires 1..200 scenes")
    if "success" in index and index["success"] is not True:
        raise ValueError("scene index must report success")
    seen_ids = set()
    assigned_item_ids = set()
    item_snapshots = {}
    scene_item_ids = []
    previous_end = 0
    total_scene_items = 0
    for scene in index["scenes"]:
        if not isinstance(scene, dict):
            raise ValueError("scene must be an object")
        start, end = _scene_range(scene)
        if start < previous_end:
            raise ValueError("scene ranges must be ordered and non-overlapping")
        previous_end = end
        items = scene.get("items")
        if not isinstance(items, list) or len(items) > MAX_SCENE_ITEMS:
            raise ValueError("scene items must be a bounded array")
        total_scene_items += len(items)
        if total_scene_items > MAX_TOTAL_SCENE_ITEMS:
            raise ValueError("scene index exceeds aggregate item budget")
        item_ids = set()
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("scene item must be an object")
            kind = item.get("type")
            if (kind is not None and
                    (not isinstance(kind, str) or not kind.strip() or
                     len(kind) > 128 or any(ord(c) < 32 or 127 <= ord(c) <= 159
                                            for c in kind))):
                raise ValueError("scene item type must be a bounded readable string")
            # Validate unknown item spans too: excluding their type from QA
            # must not allow malformed evidence to bypass integrity checks.
            if "visible_start_frame" in item or "visible_end_frame" in item:
                left, right = item.get("visible_start_frame"), item.get("visible_end_frame")
                if type(left) is not int or type(right) is not int or right <= left:
                    raise ValueError("visible item range must be valid")
            if "item_id" in item:
                identity = item["item_id"]
                if not _safe_identifier(identity, 256) or identity in item_ids:
                    raise ValueError("scene item IDs must be distinct nonempty strings")
                item_ids.add(identity)
                assigned_item_ids.add(identity)
                # Even partial source evidence must remain identical across
                # scenes: otherwise a continuation can impersonate another item.
                snapshot = tuple((key in item, item.get(key))
                                 for key in ("frame", "end_frame", "type", "layer"))
                if identity in item_snapshots and item_snapshots[identity] != snapshot:
                    raise ValueError("scene item identity has conflicting source evidence")
                item_snapshots[identity] = snapshot
            if ("layer" in item and item["layer"] is not None and
                    (type(item["layer"]) is not int or
                     not 0 <= item["layer"] <= MAX_FRAME)):
                raise ValueError("scene item layer must be a bounded integer or null")
            if "frame" in item or "end_frame" in item:
                frame, source_end = item.get("frame"), item.get("end_frame")
                if (type(frame) is not int or type(source_end) is not int or
                        frame < 0 or source_end <= frame or source_end > MAX_FRAME):
                    raise ValueError("scene source item range is invalid")
                visible_start = item.get("visible_start_frame")
                visible_end = item.get("visible_end_frame")
                start, end = _scene_range(scene)
                if max(start, frame) >= min(end, source_end):
                    raise ValueError("scene source item does not intersect its scene")
                if (visible_start != max(start, frame) or
                        visible_end != min(end, source_end) or
                        type(visible_start) is not int or type(visible_end) is not int):
                    raise ValueError("scene visible range contradicts source item")
            if "relative_start_frame" in item or "relative_end_frame" in item:
                visible_start = item.get("visible_start_frame")
                visible_end = item.get("visible_end_frame")
                relative_start = item.get("relative_start_frame")
                relative_end = item.get("relative_end_frame")
                if (type(visible_start) is not int or type(visible_end) is not int or
                        type(relative_start) is not int or type(relative_end) is not int or
                        relative_start != visible_start - start or
                        relative_end != visible_end - start or
                        not 0 <= relative_start < relative_end <= end - start):
                    raise ValueError("scene relative item range contradicts visible range")
            if "crosses_boundary" in item:
                boundary = item["crosses_boundary"]
                if (type(boundary) is not bool or
                        "frame" not in item or "end_frame" not in item or
                        boundary != (frame < start or source_end > end)):
                    raise ValueError("scene boundary flag contradicts source item")
        if "item_count" in scene and (type(scene["item_count"]) is not int or
                                       scene["item_count"] != len(items)):
            raise ValueError("scene item_count does not match item evidence")
        if "id" in scene:
            identifier = scene["id"]
            if not _safe_identifier(identifier, 64) or identifier in seen_ids:
                raise ValueError("scene IDs must be distinct nonempty strings")
            seen_ids.add(identifier)
        _validate_scene_summaries(scene, start, end)
        scene_item_ids.append(item_ids)
    # A single item may appear in multiple scenes, but it must appear in
    # *every* scene intersecting its immutable source span. Otherwise a
    # missing continuation can conceal an uncovered segment from the QA gate.
    for identity, snapshot in item_snapshots.items():
        source = {key: value for key, (present, value) in
                  zip(("frame", "end_frame", "type", "layer"), snapshot)
                  if present}
        if "frame" not in source or "end_frame" not in source:
            continue
        for scene, identities in zip(index["scenes"], scene_item_ids):
            if (source["frame"] < scene["end_frame"]
                    and source["end_frame"] > scene["start_frame"]
                    and identity not in identities):
                raise ValueError("scene item identity missing from intersecting scene")
    _validate_continuations(index["scenes"])
    if "item_count" in index or "unassigned_item_ids" in index:
        unassigned = index.get("unassigned_item_ids")
        total = index.get("item_count")
        if (type(total) is not int or not 0 <= total <= MAX_TOTAL_SCENE_ITEMS or
                not isinstance(unassigned, list) or
                len(unassigned) > MAX_TOTAL_SCENE_ITEMS or
                any(not _safe_identifier(identity, 256) for identity in unassigned) or
                len(set(unassigned)) != len(unassigned) or
                assigned_item_ids.intersection(unassigned) or
                total != len(assigned_item_ids) + len(unassigned)):
            raise ValueError("scene index item count or unassigned identities are inconsistent")
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
        # Re-evaluating a previously annotated index must not silently
        # overwrite stale or fabricated QA results. This is also idempotent
        # for a genuine second invocation.
        for key, expected in updates.items():
            if key not in scene:
                continue
            actual = scene[key]
            if key.endswith("_uncovered_ranges"):
                valid = _same_ranges(actual, expected)
            elif key.endswith("_covered_frames"):
                valid = type(actual) is int and actual == expected
            else:
                valid = type(actual) is float and actual == expected
            if not valid:
                raise ValueError(f"scene {key} contradicts item coverage")
        pending.append((scene, updates))
    # Do not leave a partly updated index when a later scene contains invalid data.
    for scene, updates in pending:
        scene.update(updates)
    return index
