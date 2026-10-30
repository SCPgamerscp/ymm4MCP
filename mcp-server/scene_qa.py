"""Opt-in structural coverage requirements for caller-defined scenes."""

from editing import integer, finite_number
import scenes
import scene_media


def _gap_statistics(scene_rows, category):
    """Find the longest missing-media interval across adjacent scene boundaries.

    Gaps between caller-defined scenes are intentionally out of scope; only
    end-exclusive intervals that touch exactly can be combined.
    """
    active = None
    longest = None
    count = 0
    top = []

    def remember(gap):
        # Bound repair hints even when a timeline contains thousands of gaps.
        top.append(gap)
        top.sort(key=lambda row: (-(row["end_frame"] - row["start_frame"]),
                                  row["start_frame"]))
        if len(top) > 10:
            top.pop()

    for scene in scene_rows:
        for gap in scene[f"{category}_uncovered_ranges"]:
            start, end = gap["start_frame"], gap["end_frame"]
            if active is not None and active["end_frame"] == start:
                active["end_frame"] = end
                if active["scene_ids"][-1] != scene["id"]:
                    active["scene_ids"].append(scene["id"])
            else:
                count += 1
                if (active is not None and
                        (longest is None or
                         active["end_frame"] - active["start_frame"] >
                         longest["end_frame"] - longest["start_frame"])):
                    longest = active
                if active is not None:
                    remember(active)
                active = {"start_frame": start, "end_frame": end,
                          "scene_ids": [scene["id"]]}
    if (active is not None and
            (longest is None or
             active["end_frame"] - active["start_frame"] >
             longest["end_frame"] - longest["start_frame"])):
        longest = active
    if active is not None:
        remember(active)
    return longest, count, top



def _covered_run_statistics(scene_rows, category, minimum):
    """Summarize continuous media occupancy, joining touching selected scenes.

    A run is a covered interval, not a count of timeline items: overlapping
    items cannot manufacture additional runs. Unselected frames never bridge
    runs, and a fully uncovered scene contributes no covered run.
    """
    active = None
    shortest = None
    run_count = 0
    short_count = 0

    def finish():
        nonlocal active, shortest, run_count, short_count
        if active is None:
            return
        run_count += 1
        duration = active["end_frame"] - active["start_frame"]
        if minimum is not None and duration < minimum:
            short_count += 1
        if (shortest is None or duration <
                shortest["end_frame"] - shortest["start_frame"]):
            shortest = active
        active = None

    for scene in scene_rows:
        cursor = scene["start_frame"]
        gaps = scene[f"{category}_uncovered_ranges"]
        # The source has already been validated by scene_media.with_coverage.
        covered = []
        for gap in gaps:
            if cursor < gap["start_frame"]:
                covered.append((cursor, gap["start_frame"]))
            cursor = gap["end_frame"]
        if cursor < scene["end_frame"]:
            covered.append((cursor, scene["end_frame"]))
        for left, right in covered:
            if active is not None and active["end_frame"] == left:
                active["end_frame"] = right
                if active["scene_ids"][-1] != scene["id"]:
                    active["scene_ids"].append(scene["id"])
            else:
                finish()
                active = {"start_frame": left, "end_frame": right,
                          "scene_ids": [scene["id"]]}
    finish()
    return shortest, run_count, short_count


def _longest_continuous_gap(scene_rows, category):
    """Compatibility helper for callers interested only in the longest gap."""
    return _gap_statistics(scene_rows, category)[0]


def inspect(items, config):
    allowed = {"scene_ranges", "require_visual", "require_audio",
               "max_visual_gap_frames", "max_audio_gap_frames",
               "max_visual_total_uncovered_frames", "max_audio_total_uncovered_frames",
               "min_visual_coverage_ratio", "min_audio_coverage_ratio",
               "max_visual_continuous_gap_frames", "max_audio_continuous_gap_frames",
               "max_visual_overall_uncovered_frames", "max_audio_overall_uncovered_frames",
               "min_visual_overall_coverage_ratio", "min_audio_overall_coverage_ratio",
               "max_visual_gap_count", "max_audio_gap_count",
               "max_visual_overall_gap_count", "max_audio_overall_gap_count",
               "max_visual_uncovered_scene_count", "max_audio_uncovered_scene_count",
               "max_visual_empty_scene_count", "max_audio_empty_scene_count",
               "min_visual_covered_run_frames", "min_audio_covered_run_frames",
               "max_visual_covered_run_count", "max_audio_covered_run_count"}
    if not isinstance(config, dict) or set(config) - allowed or "scene_ranges" not in config:
        raise ValueError("scene_check requires scene_ranges and supported coverage options")
    required, limits, total_limits, min_ratios = {}, {}, {}, {}
    continuous_limits, overall_limits, overall_min_ratios = {}, {}, {}
    gap_count_limits, overall_gap_count_limits, uncovered_scene_limits = {}, {}, {}
    empty_scene_limits = {}
    covered_run_minimums, covered_run_count_limits = {}, {}
    for category in ("visual", "audio"):
        required[category] = config.get(f"require_{category}", True)
        if not isinstance(required[category], bool):
            raise ValueError(f"require_{category} must be boolean")
        limits[category] = integer(config.get(f"max_{category}_gap_frames", 0),
                                   f"max_{category}_gap_frames")
        total_key = f"max_{category}_total_uncovered_frames"
        total_limits[category] = (integer(config[total_key], total_key)
                                  if total_key in config else None)
        ratio_key = f"min_{category}_coverage_ratio"
        min_ratios[category] = None
        if ratio_key in config:
            raw_ratio = config[ratio_key]
            if isinstance(raw_ratio, bool) or not isinstance(raw_ratio, (int, float)):
                raise ValueError(f"{ratio_key} must be a numeric ratio")
            ratio = finite_number(raw_ratio, ratio_key)
            if not 0 <= ratio <= 1:
                raise ValueError(f"{ratio_key} must be between 0 and 1")
            min_ratios[category] = ratio
        continuous_key = f"max_{category}_continuous_gap_frames"
        overall_key = f"max_{category}_overall_uncovered_frames"
        overall_ratio_key = f"min_{category}_overall_coverage_ratio"
        continuous_limits[category] = (integer(config[continuous_key], continuous_key)
                                       if continuous_key in config else None)
        overall_limits[category] = (integer(config[overall_key], overall_key)
                                    if overall_key in config else None)
        overall_min_ratios[category] = None
        if overall_ratio_key in config:
            raw = config[overall_ratio_key]
            if isinstance(raw, bool) or not isinstance(raw, (int, float)):
                raise ValueError(f"{overall_ratio_key} must be a numeric ratio")
            value = finite_number(raw, overall_ratio_key)
            if not 0 <= value <= 1:
                raise ValueError(f"{overall_ratio_key} must be between 0 and 1")
            overall_min_ratios[category] = value
        for option, destination in (
                ("gap_count", gap_count_limits),
                ("overall_gap_count", overall_gap_count_limits),
                ("uncovered_scene_count", uncovered_scene_limits),
                ("empty_scene_count", empty_scene_limits)):
            key = f"max_{category}_{option}"
            destination[category] = integer(config[key], key) if key in config else None
        run_min_key = f"min_{category}_covered_run_frames"
        run_count_key = f"max_{category}_covered_run_count"
        covered_run_minimums[category] = (
            integer(config[run_min_key], run_min_key, 1)
            if run_min_key in config else None)
        covered_run_count_limits[category] = (
            integer(config[run_count_key], run_count_key)
            if run_count_key in config else None)
        if not required[category] and any(
                key in config for key in (ratio_key, total_key, continuous_key,
                                          overall_key, overall_ratio_key,
                                           f"max_{category}_gap_count",
                                           f"max_{category}_overall_gap_count",
                                           f"max_{category}_uncovered_scene_count",
                                           f"max_{category}_empty_scene_count",
                                           run_min_key, run_count_key)):
            raise ValueError(f"disabled {category} coverage cannot have active ratio/total requirements")
    if not any(required.values()):
        raise ValueError("scene_check must require visual or audio coverage")
    index = scene_media.with_coverage(scenes.index(items, config["scene_ranges"]))
    issues = []
    for scene in index["scenes"]:
        for category in ("visual", "audio"):
            if not required[category]:
                continue
            # One issue per scene/category bounds report size even for fragmented scenes.
            all_gaps = scene[f"{category}_uncovered_ranges"]
            gaps = [gap for gap in all_gaps
                    if gap["end_frame"] - gap["start_frame"] > limits[category]]
            total_uncovered = sum(gap["end_frame"] - gap["start_frame"] for gap in all_gaps)
            total_exceeded = (total_limits[category] is not None
                              and total_uncovered > total_limits[category])
            if gaps or total_exceeded:
                # Aggregate failures must report every contributing gap, not just
                # gaps larger than the per-interval tolerance.
                reported = all_gaps if total_exceeded else gaps
                issues.append({"code": f"SCENE_{category.upper()}_GAP", "severity": "error",
                               "scene_id": scene["id"], "uncovered_ranges": reported,
                               "frame_range": [reported[0]["start_frame"], reported[-1]["end_frame"]],
                               "message": f"Scene {scene['id']} has uncovered {category} intervals"})
            gap_count_limit = gap_count_limits[category]
            if gap_count_limit is not None and len(all_gaps) > gap_count_limit:
                issues.append({
                    "code": f"SCENE_{category.upper()}_GAP_COUNT",
                    "severity": "error", "scene_id": scene["id"],
                    "gap_count": len(all_gaps), "maximum_gap_count": gap_count_limit,
                    "message": f"Scene {scene['id']} has too many uncovered {category} intervals",
                })
            threshold = min_ratios[category]
            if threshold is not None:
                length = scene["end_frame"] - scene["start_frame"]
                covered = scene[f"{category}_covered_frames"]
                if covered / length < threshold:
                    issues.append({"code": f"SCENE_{category.upper()}_COVERAGE_LOW",
                                   "severity": "error", "scene_id": scene["id"],
                                   "covered_frames": covered, "total_frames": length,
                                   "coverage_ratio": covered / length,
                                   "minimum_coverage_ratio": threshold,
                                   "message": f"Scene {scene['id']} has insufficient {category} coverage"})
    # Aggregate only the explicitly scoped scene durations. Do not include
    # unselected timeline gaps in weighted ratios or total uncovered counts.
    scoped_frames = sum(scene["end_frame"] - scene["start_frame"]
                        for scene in index["scenes"])
    overall = {}
    for category in ("visual", "audio"):
        covered = sum(scene[f"{category}_covered_frames"] for scene in index["scenes"])
        uncovered = scoped_frames - covered
        longest, gap_count, top_gaps = _gap_statistics(index["scenes"], category)
        uncovered_scene_count = sum(bool(scene[f"{category}_uncovered_ranges"])
                                    for scene in index["scenes"])
        empty_scene_count = sum(scene[f"{category}_covered_frames"] == 0
                                for scene in index["scenes"])
        ratio = covered / scoped_frames
        overall[category] = {
            "scoped_frames": scoped_frames, "covered_frames": covered,
            "uncovered_frames": uncovered, "coverage_ratio": ratio,
            "longest_uncovered_range": longest,
            "continuous_gap_count": gap_count,
            "largest_uncovered_ranges": top_gaps,
            "omitted_uncovered_range_count": gap_count - len(top_gaps),
            "uncovered_scene_count": uncovered_scene_count,
            "empty_scene_count": empty_scene_count,
        }
        if not required[category]:
            continue
        shortest_run, run_count, short_run_count = _covered_run_statistics(
            index["scenes"], category, covered_run_minimums[category])
        minimum_run = covered_run_minimums[category]
        if minimum_run is not None and short_run_count:
            issues.append({
                "code": f"SCENE_{category.upper()}_SHORT_COVERAGE_RUN",
                "severity": "error",
                "frame_range": [shortest_run["start_frame"], shortest_run["end_frame"]],
                "scene_ids": shortest_run["scene_ids"],
                "covered_frames": shortest_run["end_frame"] - shortest_run["start_frame"],
                "minimum_covered_run_frames": minimum_run,
                "short_coverage_run_count": short_run_count,
                "message": f"Short {category} coverage runs may indicate unintended flashes",
            })
        run_limit = covered_run_count_limits[category]
        if run_limit is not None and run_count > run_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_COVERED_RUN_COUNT",
                "severity": "error",
                "covered_run_count": run_count,
                "maximum_covered_run_count": run_limit,
                "message": f"Too many separate {category} coverage runs in scoped scenes",
            })
        overall_gap_limit = overall_gap_count_limits[category]
        if overall_gap_limit is not None and gap_count > overall_gap_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_OVERALL_GAP_COUNT",
                "severity": "error", "gap_count": gap_count,
                "maximum_gap_count": overall_gap_limit,
                "message": f"Too many continuous {category} gaps in scoped scenes",
            })
        scene_limit = uncovered_scene_limits[category]
        if scene_limit is not None and uncovered_scene_count > scene_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_UNCOVERED_SCENE_COUNT",
                "severity": "error", "uncovered_scene_count": uncovered_scene_count,
                "maximum_uncovered_scene_count": scene_limit,
                "message": f"Too many scenes have uncovered {category} frames",
            })
        empty_limit = empty_scene_limits[category]
        if empty_limit is not None and empty_scene_count > empty_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_EMPTY_SCENE_COUNT",
                "severity": "error", "empty_scene_count": empty_scene_count,
                "maximum_empty_scene_count": empty_limit,
                "message": f"Too many scenes lack any {category} coverage",
            })
        continuous_limit = continuous_limits[category]
        if (continuous_limit is not None and longest is not None and
                longest["end_frame"] - longest["start_frame"] > continuous_limit):
            issues.append({
                "code": f"SCENE_{category.upper()}_CONTINUOUS_GAP",
                "severity": "error",
                "frame_range": [longest["start_frame"], longest["end_frame"]],
                "scene_ids": longest["scene_ids"],
                "gap_frames": longest["end_frame"] - longest["start_frame"],
                "maximum_gap_frames": continuous_limit,
                "message": f"Continuous {category} gap exceeds scoped limit",
            })
        overall_limit = overall_limits[category]
        if overall_limit is not None and uncovered > overall_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_OVERALL_UNCOVERED",
                "severity": "error", "uncovered_frames": uncovered,
                "scoped_frames": scoped_frames, "maximum_uncovered_frames": overall_limit,
                "message": f"Total uncovered {category} frames exceed scoped limit",
            })
        minimum = overall_min_ratios[category]
        if minimum is not None and ratio < minimum:
            issues.append({
                "code": f"SCENE_{category.upper()}_OVERALL_COVERAGE_LOW",
                "severity": "error", "covered_frames": covered,
                "scoped_frames": scoped_frames, "coverage_ratio": ratio,
                "minimum_coverage_ratio": minimum,
                "message": f"Overall {category} coverage is below scoped minimum",
            })
    return {"success": True, "passed": not issues, "issues": issues, "scenes": index["scenes"],
            "overall_coverage": overall,
            "scope": "Timeline occupancy only; pixels, opacity, volume and embedded video audio are not inspected."}
