"""Opt-in structural coverage requirements for caller-defined scenes."""

from editing import integer
import scenes
import scene_media


def inspect(items, config):
    allowed = {"scene_ranges", "require_visual", "require_audio",
               "max_visual_gap_frames", "max_audio_gap_frames"}
    if not isinstance(config, dict) or set(config) - allowed or "scene_ranges" not in config:
        raise ValueError("scene_check requires scene_ranges and supported coverage options")
    required, limits = {}, {}
    for category in ("visual", "audio"):
        required[category] = config.get(f"require_{category}", True)
        if not isinstance(required[category], bool):
            raise ValueError(f"require_{category} must be boolean")
        limits[category] = integer(config.get(f"max_{category}_gap_frames", 0),
                                   f"max_{category}_gap_frames")
    if not any(required.values()):
        raise ValueError("scene_check must require visual or audio coverage")
    index = scene_media.with_coverage(scenes.index(items, config["scene_ranges"]))
    issues = []
    for scene in index["scenes"]:
        for category in ("visual", "audio"):
            if not required[category]:
                continue
            # One issue per scene/category bounds report size even for fragmented scenes.
            gaps = [gap for gap in scene[f"{category}_uncovered_ranges"]
                    if gap["end_frame"] - gap["start_frame"] > limits[category]]
            if gaps:
                issues.append({"code": f"SCENE_{category.upper()}_GAP", "severity": "error",
                               "scene_id": scene["id"], "uncovered_ranges": gaps,
                               "frame_range": [gaps[0]["start_frame"], gaps[-1]["end_frame"]],
                               "message": f"Scene {scene['id']} has uncovered {category} intervals"})
    return {"success": True, "passed": not issues, "issues": issues, "scenes": index["scenes"],
            "scope": "Timeline occupancy only; pixels, opacity, volume and embedded video audio are not inspected."}
