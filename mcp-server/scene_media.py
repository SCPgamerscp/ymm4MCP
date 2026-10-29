"""Estimate independent visual and audio coverage from a read-only scene index."""


def is_known_item_type(value: object, names: tuple[str, ...]) -> bool:
    """Match the exact CLR type name, optionally with namespace/nesting qualifiers."""
    if not isinstance(value, str):
        return False
    leaf = value.rsplit(".", 1)[-1].rsplit("+", 1)[-1]
    return leaf.lower() in names


def _scene_range(scene: dict) -> tuple[int, int]:
    start, end = scene["start_frame"], scene["end_frame"]
    if type(start) is not int or type(end) is not int or end <= start:
        raise ValueError("scene range must have integer frames and positive duration")
    return start, end


def _spans(scene: dict, names: tuple[str, ...], start: int, end: int) -> list[tuple[int, int]]:
    spans = []
    for item in scene["items"]:
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
