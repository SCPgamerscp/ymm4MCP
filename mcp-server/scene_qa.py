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
    longest = None
    run_count = 0
    short_count = 0
    short_frames = 0

    def finish():
        nonlocal active, shortest, longest, run_count, short_count, short_frames
        if active is None:
            return
        run_count += 1
        duration = active["end_frame"] - active["start_frame"]
        if minimum is not None and duration < minimum:
            short_count += 1
            short_frames += duration
        if (shortest is None or duration <
                shortest["end_frame"] - shortest["start_frame"]):
            shortest = active
        if (longest is None or duration >
                longest["end_frame"] - longest["start_frame"]):
            longest = active
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
    return shortest, longest, run_count, short_count, short_frames



def _edge_uncovered_frames(scene_rows, category, *, trailing=False):
    """Count missing coverage at a selected-scope edge, not across excluded time."""
    previous_boundary = None
    total = 0
    rows = reversed(scene_rows) if trailing else iter(scene_rows)
    for scene in rows:
        start, end = scene["start_frame"], scene["end_frame"]
        if previous_boundary is not None:
            if (end if trailing else start) != previous_boundary:
                break
        gaps = scene[f"{category}_uncovered_ranges"]
        edge = (gaps[-1] if trailing else gaps[0]) if gaps else None
        if edge is None or (edge["end_frame"] != end if trailing
                            else edge["start_frame"] != start):
            break
        total += edge["end_frame"] - edge["start_frame"]
        if edge["start_frame"] != start or edge["end_frame"] != end:
            break
        previous_boundary = start if trailing else end
    return total


def _empty_scene_streaks(scene_rows, category):
    """Find longest touching fully empty scenes by count and by duration."""
    current = None
    by_count = None
    by_frames = None

    def finish():
        nonlocal current, by_count, by_frames
        if current is None:
            return
        if by_count is None or current["scene_count"] > by_count["scene_count"]:
            by_count = current
        if by_frames is None or current["frames"] > by_frames["frames"]:
            by_frames = current
        current = None

    for scene in scene_rows:
        start, end = scene["start_frame"], scene["end_frame"]
        if scene[f"{category}_covered_frames"] != 0:
            finish()
            continue
        if current is None or current["end_frame"] != start:
            finish()
            current = {"start_frame": start, "end_frame": end,
                       "scene_ids": [scene["id"]], "scene_count": 1,
                       "frames": end - start}
        else:
            current["end_frame"] = end
            current["scene_ids"].append(scene["id"])
            current["scene_count"] += 1
            current["frames"] += end - start
    finish()
    return by_count, by_frames


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
               "max_visual_covered_run_count", "max_audio_covered_run_count",
               "max_visual_short_covered_run_count", "max_audio_short_covered_run_count",
               "min_visual_covered_run_count", "min_audio_covered_run_count",
               "min_visual_longest_covered_run_frames", "min_audio_longest_covered_run_frames",
               "max_visual_short_covered_run_total_frames", "max_audio_short_covered_run_total_frames",
               "max_visual_leading_uncovered_frames", "max_audio_leading_uncovered_frames",
               "max_visual_trailing_uncovered_frames", "max_audio_trailing_uncovered_frames",
               "max_visual_consecutive_empty_scene_count", "max_audio_consecutive_empty_scene_count",
               "max_visual_consecutive_empty_scene_frames", "max_audio_consecutive_empty_scene_frames"}
    if not isinstance(config, dict) or set(config) - allowed or "scene_ranges" not in config:
        raise ValueError("scene_check requires scene_ranges and supported coverage options")
    required, limits, total_limits, min_ratios = {}, {}, {}, {}
    continuous_limits, overall_limits, overall_min_ratios = {}, {}, {}
    gap_count_limits, overall_gap_count_limits, uncovered_scene_limits = {}, {}, {}
    empty_scene_limits = {}
    leading_limits, trailing_limits = {}, {}
    empty_streak_count_limits, empty_streak_frame_limits = {}, {}
    covered_run_minimums, covered_run_count_limits = {}, {}
    short_run_count_limits = {}
    minimum_run_counts, minimum_longest_runs, short_run_frame_limits = {}, {}, {}
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
                ("empty_scene_count", empty_scene_limits),
                ("leading_uncovered_frames", leading_limits),
                ("trailing_uncovered_frames", trailing_limits),
                ("consecutive_empty_scene_count", empty_streak_count_limits),
                ("consecutive_empty_scene_frames", empty_streak_frame_limits)):
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
        short_count_key = f"max_{category}_short_covered_run_count"
        short_run_count_limits[category] = (
            integer(config[short_count_key], short_count_key)
            if short_count_key in config else 0)
        if short_count_key in config and run_min_key not in config:
            raise ValueError(f"{short_count_key} requires {run_min_key}")
        min_count_key = f"min_{category}_covered_run_count"
        longest_key = f"min_{category}_longest_covered_run_frames"
        short_frames_key = f"max_{category}_short_covered_run_total_frames"
        minimum_run_counts[category] = (
            integer(config[min_count_key], min_count_key, 1)
            if min_count_key in config else None)
        minimum_longest_runs[category] = (
            integer(config[longest_key], longest_key, 1)
            if longest_key in config else None)
        short_run_frame_limits[category] = (
            integer(config[short_frames_key], short_frames_key)
            if short_frames_key in config else None)
        if short_frames_key in config and run_min_key not in config:
            raise ValueError(f"{short_frames_key} requires {run_min_key}")
        if not required[category] and any(
                key in config for key in (ratio_key, total_key, continuous_key,
                                          overall_key, overall_ratio_key,
                                           f"max_{category}_gap_count",
                                           f"max_{category}_overall_gap_count",
                                           f"max_{category}_uncovered_scene_count",
                                           f"max_{category}_empty_scene_count",
                                            f"max_{category}_leading_uncovered_frames",
                                            f"max_{category}_trailing_uncovered_frames",
                                            f"max_{category}_consecutive_empty_scene_count",
                                            f"max_{category}_consecutive_empty_scene_frames",
                                           run_min_key, run_count_key, short_count_key,
                                           min_count_key, longest_key, short_frames_key)):
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
        for edge_name, edge_limits, trailing in (
                ("leading", leading_limits, False),
                ("trailing", trailing_limits, True)):
            limit = edge_limits[category]
            if limit is None:
                continue
            missing = _edge_uncovered_frames(index["scenes"], category,
                                             trailing=trailing)
            overall[category][f"{edge_name}_uncovered_frames"] = missing
            if missing > limit:
                issues.append({
                    "code": f"SCENE_{category.upper()}_{edge_name.upper()}_UNCOVERED",
                    "severity": "error", "uncovered_frames": missing,
                    "maximum_uncovered_frames": limit,
                    "message": f"Selected scenes have excessive {edge_name} {category} gaps",
                })
        count_limit = empty_streak_count_limits[category]
        frame_limit = empty_streak_frame_limits[category]
        if count_limit is not None or frame_limit is not None:
            by_count, by_frames = _empty_scene_streaks(index["scenes"], category)
            overall[category]["longest_empty_scene_streak_by_count"] = by_count
            overall[category]["longest_empty_scene_streak_by_frames"] = by_frames
            if count_limit is not None and by_count is not None and by_count["scene_count"] > count_limit:
                issues.append({
                    "code": f"SCENE_{category.upper()}_CONSECUTIVE_EMPTY_SCENES",
                    "severity": "error", "scene_ids": by_count["scene_ids"],
                    "frame_range": [by_count["start_frame"], by_count["end_frame"]],
                    "empty_scene_count": by_count["scene_count"],
                    "maximum_empty_scene_count": count_limit,
                    "message": f"Too many consecutive scenes lack {category} coverage",
                })
            if frame_limit is not None and by_frames is not None and by_frames["frames"] > frame_limit:
                issues.append({
                    "code": f"SCENE_{category.upper()}_CONSECUTIVE_EMPTY_FRAMES",
                    "severity": "error", "scene_ids": by_frames["scene_ids"],
                    "frame_range": [by_frames["start_frame"], by_frames["end_frame"]],
                    "empty_frames": by_frames["frames"],
                    "maximum_empty_frames": frame_limit,
                    "message": f"Consecutive empty {category} scenes exceed frame budget",
                })
        shortest_run, longest_run, run_count, short_run_count, short_frames = (
            _covered_run_statistics(index["scenes"], category,
                                    covered_run_minimums[category]))
        minimum_run = covered_run_minimums[category]
        # Emit bounded evidence when a covered-run gate is requested, so
        # automated repair can locate the shortest visible/audible segment.
        if (minimum_run is not None or covered_run_count_limits[category] is not None
                or minimum_run_counts[category] is not None
                or minimum_longest_runs[category] is not None):
            overall[category].update({
                "covered_run_count": run_count,
                "short_covered_run_count": short_run_count,
                "short_covered_frames": short_frames,
                "shortest_covered_range": shortest_run,
                "longest_covered_range": longest_run,
            })
        if (minimum_run is not None and
                short_run_count > short_run_count_limits[category]):
            issues.append({
                "code": f"SCENE_{category.upper()}_SHORT_COVERAGE_RUN",
                "severity": "error",
                "frame_range": [shortest_run["start_frame"], shortest_run["end_frame"]],
                "scene_ids": shortest_run["scene_ids"],
                "covered_frames": shortest_run["end_frame"] - shortest_run["start_frame"],
                "minimum_covered_run_frames": minimum_run,
                "short_coverage_run_count": short_run_count,
                "maximum_short_coverage_run_count": short_run_count_limits[category],
                "message": f"Short {category} coverage runs may indicate unintended flashes",
            })
        short_frames_limit = short_run_frame_limits[category]
        if short_frames_limit is not None and short_frames > short_frames_limit:
            issues.append({
                "code": f"SCENE_{category.upper()}_SHORT_COVERAGE_FRAMES",
                "severity": "error", "short_covered_frames": short_frames,
                "maximum_short_covered_frames": short_frames_limit,
                "message": f"Total short {category} coverage frames exceed budget",
            })
        min_run_count = minimum_run_counts[category]
        if min_run_count is not None and run_count < min_run_count:
            issues.append({
                "code": f"SCENE_{category.upper()}_COVERED_RUN_COUNT_LOW",
                "severity": "error", "covered_run_count": run_count,
                "minimum_covered_run_count": min_run_count,
                "message": f"Too few separate {category} coverage runs in scoped scenes",
            })
        min_longest = minimum_longest_runs[category]
        longest_frames = (longest_run["end_frame"] - longest_run["start_frame"]
                          if longest_run is not None else 0)
        if min_longest is not None and longest_frames < min_longest:
            issues.append({
                "code": f"SCENE_{category.upper()}_LONGEST_COVERAGE_RUN_LOW",
                "severity": "error", "longest_covered_frames": longest_frames,
                "minimum_longest_covered_run_frames": min_longest,
                "longest_covered_range": longest_run,
                "message": f"No sustained {category} coverage run meets the minimum",
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
