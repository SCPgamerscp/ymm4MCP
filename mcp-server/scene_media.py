"""Estimate independent visual and audio coverage from a read-only scene index."""


def is_known_item_type(value: object, names: tuple[str, ...]) -> bool:
    """Match the exact CLR type name, optionally with namespace/nesting qualifiers."""
    if not isinstance(value, str):
        return False
    leaf = value.rsplit(".", 1)[-1].rsplit("+", 1)[-1]
    return leaf.lower() in names


def with_coverage(index: dict) -> dict:
    # Unknown host item types are excluded rather than assumed to provide a picture.
    categories = {"visual": ("videoitem", "imageitem", "textitem", "tachieitem",
                             "tachiefaceitem", "faceitem"),
                  "audio": ("audioitem", "voiceitem")}
    for scene in index["scenes"]:
        for category, suffixes in categories.items():
            spans = sorted((item["visible_start_frame"], item["visible_end_frame"])
                           for item in scene["items"]
                           if is_known_item_type(item["type"], suffixes))
            cursor = scene["start_frame"]
            gaps = []
            for left, right in spans:
                if left > cursor:
                    gaps.append({"start_frame": cursor, "end_frame": left})
                cursor = max(cursor, right)
            if cursor < scene["end_frame"]:
                gaps.append({"start_frame": cursor, "end_frame": scene["end_frame"]})
            scene[f"{category}_uncovered_ranges"] = gaps
            scene[f"{category}_covered_frames"] = (scene["end_frame"] - scene["start_frame"] -
                                                    sum(g["end_frame"] - g["start_frame"] for g in gaps))
    return index
