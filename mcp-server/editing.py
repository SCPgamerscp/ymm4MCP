"""Side-effect-free planning and verification for timeline edits."""
import hashlib
import json
import math
from collections import Counter

MAX_FRAME = 2_147_483_647
MAX_QA_ISSUES = 1000


def _bounded_qa_issues(issues):
    """Keep QA gate reports usable without silently accepting truncated failures."""
    if len(issues) <= MAX_QA_ISSUES:
        return issues
    # A dropped truncation marker represents its underlying omitted issues,
    # not one additional issue. Preserve nested source-report omissions when
    # combining already-bounded structural, visual and audio evidence.
    omitted = sum(issue["omitted_count"] if issue.get("code") == "QA_ISSUES_TRUNCATED" else 1
                  for issue in issues[MAX_QA_ISSUES - 1:])
    return issues[:MAX_QA_ISSUES - 1] + [{
        "code": "QA_ISSUES_TRUNCATED", "severity": "error",
        "omitted_count": omitted,
        "message": f"{omitted} additional QA issues were omitted; inspect the source reports",
    }]


def _require_json_safe_issues(issues, source):
    """Reject non-JSON and non-finite evidence before storing QA history."""
    try:
        json.dumps(issues, ensure_ascii=False, allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError, UnicodeError) as exc:
        raise ValueError(f"{source} issues must contain JSON-safe finite UTF-8 values") from exc




def _valid_qa_code(value):
    """Machine-readable QA codes must be bounded printable ASCII tokens."""
    return (isinstance(value, str) and 1 <= len(value) <= 128 and
            all(33 <= ord(char) <= 126 for char in value))


def integer(value, name, minimum=0, maximum=MAX_FRAME):
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise ValueError(f"{name} must be an integer in {minimum}..{maximum}")
    return value


def finite_number(value, name):
    if isinstance(value, bool):
        raise ValueError(f"{name} must be a finite number")
    if isinstance(value, str):
        try:
            value = float(value)
        except ValueError as exc:
            raise ValueError(f"{name} must be a finite number") from exc
    if not isinstance(value, (int, float)):
        raise ValueError(f"{name} must be a finite number")
    try:
        if not math.isfinite(value):
            raise ValueError(f"{name} must be a finite number")
        return float(value)
    except OverflowError as exc:
        raise ValueError(f"{name} must be a finite number") from exc


def plan_script(args):
    lines = args.get("lines", [])
    if not isinstance(lines, list) or not 1 <= len(lines) <= 500:
        raise ValueError("lines must contain 1..500 dialogue entries")
    fps = integer(args.get("fps", 30), "fps", 1, 240)
    speed = args.get("chars_per_sec", 5)
    if isinstance(speed, bool) or not isinstance(speed, (int, float)) or not 0 < speed <= 1000 or not math.isfinite(speed):
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


def _valid_item_id(value):
    """Only usable host IDs may be referenced in repair suggestions."""
    return (isinstance(value, str) and bool(value.strip()) and
            not any(ord(char) < 32 or ord(char) == 127 or
                    0xD800 <= ord(char) <= 0xDFFF for char in value))


def _item_ids(items, indices):
    """Return only usable IDs present on the affected items."""
    return [items[index]["item_id"] for index in indices
            if isinstance(items[index], dict) and _valid_item_id(items[index].get("item_id"))]


def validate_timeline(items, expected=None, duration=None, include_gaps=True, subtitle_layers=None,
                      subtitle_timing_tolerance_frames=0):
    """Produce machine-readable structural QA without changing the timeline."""
    if not isinstance(items, list):
        raise ValueError("items must be an array")
    if duration is not None:
        integer(duration, "duration", 1)
    if not isinstance(include_gaps, bool):
        raise ValueError("include_gaps must be boolean")
    tolerance = integer(subtitle_timing_tolerance_frames, "subtitle_timing_tolerance_frames")
    if subtitle_layers is not None:
        if not isinstance(subtitle_layers, list) or not 1 <= len(subtitle_layers) <= 128:
            raise ValueError("subtitle_layers must contain 1..128 layer numbers")
        subtitle_layers = {integer(layer, "subtitle_layers entry") for layer in subtitle_layers}

    problems = []
    layers = {}
    voices = []
    subtitles = []
    seen_item_ids = {}
    valid_indices = set()
    # Count all stable IDs before offering an edit: duplicates later in the
    # snapshot are just as ambiguous as duplicates already visited.
    id_counts = Counter(item["item_id"] for item in items
                        if isinstance(item, dict) and _valid_item_id(item.get("item_id")))
    for index, item in enumerate(items):
        try:
            if not isinstance(item, dict):
                raise ValueError("item must be an object")
            frame = integer(item.get("frame"), "frame")
            length = integer(item.get("length"), "length", 1)
            layer = integer(item.get("layer"), "layer")
            end = integer(frame + length, "end frame")
        except ValueError as exc:
            ids = _item_ids(items, [index])
            problems.append({"code": "INVALID_ITEM", "severity": "error", "index": index,
                             "item_ids": ids, "message": str(exc)})
            continue
        valid_indices.add(index)
        # A repeated stable item_id makes follow-up edits and QA repairs target
        # an arbitrary timeline item. Missing IDs are permitted for legacy
        # snapshots, but any supplied non-null ID must be usable and unique.
        if item.get("item_id") is not None:
            item_id = item["item_id"]
            if not _valid_item_id(item_id):
                problems.append({"code": "INVALID_ITEM_ID", "severity": "error",
                                 "index": index, "item_ids": [],
                                 "message": "Timeline item has an invalid stable item_id"})
            elif item_id in seen_item_ids:
                previous = seen_item_ids[item_id]
                problems.append({"code": "DUPLICATE_ITEM_ID", "severity": "error",
                                 "indices": [previous, index], "item_ids": [item_id],
                                 "message": "Multiple timeline items share the same item_id"})
            else:
                seen_item_ids[item_id] = index
        layers.setdefault(layer, []).append((frame, end, index))
        if subtitle_layers is not None:
            kind = item.get("type")
            # Compare the actual CLR type name, not an arbitrary suffix such as
            # NotVoiceItem. Preserve support for namespace and nested type names.
            kind = kind.replace("+", ".").rsplit(".", 1)[-1].casefold() if isinstance(kind, str) else ""
            if kind in ("voiceitem", "voice"):
                voices.append((frame, end, index))
            elif layer in subtitle_layers and kind in ("textitem", "text"):
                subtitles.append((frame, end, index))
        if duration is not None and end > duration:
            problem = {"code": "EXCEEDS_DURATION", "severity": "error", "index": index,
                       "item_ids": _item_ids(items, [index]), "frame_range": [frame, end],
                       "project_duration": duration}
            if frame < duration:
                # A repair hint must identify exactly one real item. An absent,
                # malformed or duplicated ID could edit the wrong item.
                item_id = item.get("item_id")
                if _valid_item_id(item_id) and id_counts[item_id] == 1:
                    problem["suggested_fix"] = {"action": "edit_item", "sub_action": "property",
                                                "item_id": item_id,
                                                "prop": "Length", "value": duration - frame}
                else:
                    problem["message"] = "Cannot safely trim an item without a unique stable item_id"
            else:
                # Moving to duration-1 leaves any item longer than one frame
                # out of bounds. Never propose a non-fixing automatic edit.
                problem["message"] = "Item starts outside the project duration; review its placement"
            problems.append(problem)

    for layer, spans in layers.items():
        active_end = None
        active_index = None
        for frame, tail, index in sorted(spans):
            if active_index is not None and frame < active_end:
                affected = [active_index, index]
                problems.append({"code": "OVERLAP", "severity": "error", "layer": layer,
                                 "indices": affected, "item_ids": _item_ids(items, affected),
                                 "frame_range": [frame, min(active_end, tail)],
                                 "suggested_fix": {"action": "edit_item", "sub_action": "resolve_overlaps",
                                                   "layers": [layer]}})
            elif include_gaps and active_index is not None and frame > active_end:
                problems.append({"code": "GAP", "severity": "warning", "layer": layer,
                                 "indices": [active_index, index],
                                 "item_ids": _item_ids(items, [active_index, index]),
                                 "frame_range": [active_end, frame]})
            if active_end is None or tail > active_end:
                active_end, active_index = tail, index

    if subtitle_layers is not None:
        for frame, end, index in voices:
            voice_text = items[index].get("text")
            if not isinstance(voice_text, str) or not voice_text.strip():
                problems.append({"code": "SUBTITLE_CHECK_SKIPPED", "severity": "warning",
                                 "index": index, "item_ids": _item_ids(items, [index]),
                                 "frame_range": [frame, end], "message": "発話テキストを取得できません"})
                continue
            normalized = "".join(voice_text.split())
            overlapping = [(s_index, items[s_index].get("text")) for start, stop, s_index in subtitles
                           if start < end and frame < stop]
            matching = [(max(start, frame), min(stop, end), s_index)
                        for start, stop, s_index in subtitles
                        if start < end and frame < stop and
                        isinstance(items[s_index].get("text"), str) and
                        "".join(items[s_index]["text"].split()) == normalized]
            if matching:
                cursor, gaps = frame, []
                for left, right, _ in sorted(matching):
                    if left - cursor > tolerance:
                        gaps.append({"start_frame": cursor, "end_frame": left})
                    cursor = max(cursor, right)
                if end - cursor > tolerance:
                    gaps.append({"start_frame": cursor, "end_frame": end})
                if gaps:
                    affected = [index, *(s_index for _, _, s_index in matching)]
                    problems.append({"code": "SUBTITLE_TIMING_MISMATCH", "severity": "error",
                                     "index": index, "indices": affected,
                                     "item_ids": _item_ids(items, affected),
                                     "frame_range": [frame, end], "uncovered_ranges": gaps,
                                     "expected_text": voice_text})
                # A correct subtitle must not mask a second, conflicting
                # subtitle on another configured subtitle layer.
                conflicting = [(s_index, text) for s_index, text in overlapping
                               if not isinstance(text, str) or "".join(text.split()) != normalized]
                if conflicting:
                    affected = [index, *(s_index for s_index, _ in conflicting)]
                    problems.append({"code": "SUBTITLE_TEXT_MISMATCH", "severity": "error",
                                     "index": index, "indices": affected,
                                     "item_ids": _item_ids(items, affected),
                                     "frame_range": [frame, end], "expected_text": voice_text,
                                     "actual_texts": [text for _, text in conflicting]})
                continue
            matches = [s_index for s_index, _ in overlapping]
            problems.append({
                "code": "SUBTITLE_TEXT_MISMATCH" if matches else "SUBTITLE_MISSING",
                "severity": "error", "index": index, "indices": [index, *matches],
                "item_ids": _item_ids(items, [index, *matches]),
                "frame_range": [frame, end], "expected_text": voice_text,
                "actual_texts": [text for _, text in overlapping],
            })

    if expected is not None:
        if not isinstance(expected, list) or len(expected) > 1000:
            raise ValueError("expected must be an array of at most 1000 item descriptions")
        supported = {"frame", "layer", "length", "type", "text", "item_id", "revision", "id"}
        claimed_item_indices = set()
        for index, wanted in enumerate(expected):
            if not isinstance(wanted, dict) or not wanted or set(wanted) - supported:
                raise ValueError("expected entries must contain supported nonempty item selectors")
            # Python considers True == 1 and None == a missing dict value.
            # Validate selectors before matching to avoid false QA passes.
            for key, value in wanted.items():
                if key in {"frame", "layer", "length"}:
                    integer(value, f"expected[{index}].{key}", 1 if key == "length" else 0)
                elif not isinstance(value, str) or (key != "text" and not value.strip()) or (
                        key in ("id", "item_id") and not _valid_item_id(value)):
                    raise ValueError(f"expected[{index}].{key} must be a valid string")
                if isinstance(value, str):
                    try:
                        value.encode("utf-8")
                    except UnicodeError as exc:
                        raise ValueError(f"expected[{index}].{key} must be valid UTF-8") from exc
            selector = dict(wanted)
            if "id" in selector:
                legacy_id = selector.pop("id")
                if "item_id" in selector and selector["item_id"] != legacy_id:
                    raise ValueError("expected id and item_id must not conflict")
                selector["item_id"] = legacy_id
            matches = [i for i, item in enumerate(items)
                       if i in valid_indices and isinstance(item, dict) and
                       all(item.get(k) == v for k, v in selector.items())]
            if len(matches) == 1 and matches[0] in claimed_item_indices:
                problems.append({"code": "EXPECTED_REUSED", "severity": "error",
                                 "expected_index": index, "expected": selector,
                                 "matches": matches, "item_ids": _item_ids(items, matches),
                                 "message": "Multiple expected entries resolve to the same timeline item"})
            elif len(matches) == 1:
                claimed_item_indices.add(matches[0])
            else:
                problems.append({"code": "EXPECTED_NOT_FOUND" if not matches else "EXPECTED_AMBIGUOUS",
                                 "severity": "error", "expected_index": index,
                                 "expected": selector, "matches": matches,
                                 "item_ids": _item_ids(items, matches)})

    problems = _bounded_qa_issues(problems)
    error_count = sum(problem["severity"] == "error" for problem in problems)
    warning_count = sum(problem["severity"] == "warning" for problem in problems)
    criteria = {"expected": expected, "duration": duration, "include_gaps": include_gaps,
                "subtitle_layers": sorted(subtitle_layers) if subtitle_layers is not None else None,
                "subtitle_timing_tolerance_frames": tolerance}
    criteria_hash = hashlib.sha256(json.dumps(criteria, sort_keys=True, ensure_ascii=False,
                                             separators=(",", ":"), default=str).encode("utf-8")).hexdigest()
    return {"success": True, "passed": error_count == 0, "valid": error_count == 0,
            "criteria_hash": criteria_hash,
            "score": max(0, 100 - error_count * 20 - warning_count * 5),
            "item_count": len(items), "issue_count": len(problems),
            "summary": {"errors": error_count, "warnings": warning_count},
            "issues": problems, "problems": problems,
            "scope": "Frame ranges, same-layer overlaps/internal gaps, explicit expected items, and optional voice/subtitle text and interval coverage after whitespace removal; visual/audio quality is not checked."}


def evaluate_qa_gate(current, history=None, *, max_repairs=3, repeat_limit=2,
                     elapsed_seconds=0, max_seconds=None, api_calls=0, max_api_calls=None):
    """Decide whether to accept, repair, or stop after a read-only timeline QA run.

    History contains prior validate results in chronological order. Each result must
    have been produced with the same validation options as the current result.
    Counters are supplied by the caller; this function never edits or rolls back.
    """
    if history is None:
        history = []
    if not isinstance(history, list) or len(history) > 20:
        raise ValueError("qa_history must contain at most 20 prior reports")
    max_repairs = integer(max_repairs, "max_repairs", 0, 20)
    repeat_limit = integer(repeat_limit, "repeat_limit", 2, 10)
    elapsed_seconds = finite_number(elapsed_seconds, "elapsed_seconds")
    api_calls = integer(api_calls, "api_calls")
    if elapsed_seconds < 0:
        raise ValueError("elapsed_seconds must be non-negative")
    if max_seconds is not None:
        max_seconds = finite_number(max_seconds, "max_seconds")
        if max_seconds <= 0:
            raise ValueError("max_seconds must be positive")
    if max_api_calls is not None:
        integer(max_api_calls, "max_api_calls", 1)

    def check_report(report):
        if not isinstance(report, dict) or report.get("success") is not True or not isinstance(report.get("passed"), bool):
            raise ValueError("qa_history and current must contain successful validate reports")
        score = finite_number(report.get("score"), "score")
        if not 0 <= score <= 100 or not isinstance(report.get("issues"), list) or len(report["issues"]) > 1000:
            raise ValueError("validate reports require a 0..100 score and at most 1000 issues")
        _require_json_safe_issues(report["issues"], "validate report")
        has_errors = False
        errors = warnings = 0
        for issue in report["issues"]:
            if not isinstance(issue, dict) or not _valid_qa_code(issue.get("code")):
                raise ValueError("validate report issues require a code")
            if issue.get("severity") not in ("error", "warning"):
                raise ValueError("validate report issues require error or warning severity")
            if issue["code"] == "QA_ISSUES_TRUNCATED" and (
                    issue["severity"] != "error" or
                    type(issue.get("omitted_count")) is not int or
                    issue["omitted_count"] <= 0):
                raise ValueError("QA_ISSUES_TRUNCATED requires a positive omitted_count and error severity")
            if "item_ids" in issue and (not isinstance(issue["item_ids"], list) or
                    any(not _valid_item_id(identity) for identity in issue["item_ids"])):
                raise ValueError("validate report item_ids must be nonempty strings")
            if issue["severity"] == "error":
                errors += 1
                has_errors = True
            else:
                warnings += 1
        if report["passed"] == has_errors:
            raise ValueError("validate report passed status disagrees with error issues")
        if score != max(0, 100 - errors * 20 - warnings * 5):
            raise ValueError("validate report score disagrees with issue severities")
        if "valid" in report and report["valid"] is not report["passed"]:
            raise ValueError("validate report valid status disagrees with passed")
        if "issue_count" in report and (type(report["issue_count"]) is not int or
                                        report["issue_count"] != len(report["issues"])):
            raise ValueError("validate report issue_count disagrees with issues")
        if "item_count" in report and (type(report["item_count"]) is not int or
                                       report["item_count"] < 0):
            raise ValueError("validate report item_count must be a non-negative integer")
        if "problems" in report and report["problems"] != report["issues"]:
            raise ValueError("validate report problems disagrees with issues")
        if "summary" in report:
            summary = report["summary"]
            if (not isinstance(summary, dict) or
                    type(summary.get("errors")) is not int or
                    type(summary.get("warnings")) is not int or
                    summary["errors"] != errors or summary["warnings"] != warnings):
                raise ValueError("validate report summary disagrees with issue severities")
        # The bounded report retains only the first 999 issues and a marker.
        # Preserve the hidden count so worsening quality at the score floor is visible.
        omitted = sum(issue.get("omitted_count", 0) for issue in report["issues"]
                      if issue["code"] == "QA_ISSUES_TRUNCATED")
        return score, errors, warnings, omitted

    metrics = [check_report(report) for report in [*history, current]]
    scores = [metric[0] for metric in metrics]
    if not isinstance(current.get("criteria_hash"), str) or not current["criteria_hash"].strip() or any(
            report.get("criteria_hash") != current["criteria_hash"] for report in history):
        raise ValueError("qa_history must use the same validation criteria as current QA")

    def signature(report):
        # Compare complete issue evidence, not a hand-picked list of fields.
        # Positional indices change when the host reorders the timeline; stable
        # item IDs and actual QA evidence do not. Keep unknown future fields so
        # a meaningful repair cannot be mistaken for an unchanged QA result.
        def canonical(issue):
            evidence = {key: value for key, value in issue.items()
                        if key not in ("index", "indices", "matches")}
            # Timeline positions are unstable, but the number of selector
            # matches is meaningful: resolving one of three anonymous matches
            # must count as progress even while EXPECTED_AMBIGUOUS remains.
            if "matches" in issue:
                matches = issue["matches"]
                evidence["match_count"] = len(matches) if isinstance(matches, list) else matches
            if "item_ids" in evidence:
                evidence["item_ids"] = sorted(evidence["item_ids"])
            return json.dumps(evidence, sort_keys=True, ensure_ascii=False,
                              allow_nan=False, separators=(",", ":"))

        # Identical diagnostics can describe different snapshots when the
        # number of timeline items changes. That is progress, not a stalled edit.
        return report.get("item_count"), sorted(canonical(issue) for issue in report["issues"])

    def error_identity(issue):
        """Track new hard failures even when other repairs raise the QA score.

        A stable item ID distinguishes separate failures with the same code.
        When the host cannot provide IDs, use code/source rather than unstable
        timeline indices; frame moves alone must not imply a new failure.
        """
        ids = issue.get("item_ids") or []
        return json.dumps([issue["code"], issue.get("source"), sorted(ids)],
                          sort_keys=True, ensure_ascii=False, allow_nan=False,
                          separators=(",", ":"))

    def introduces_new_error(previous, latest):
        # Compare multiplicities too: three identical black-frame failures
        # after two is a new failure even if another error was repaired.
        previous_errors = Counter(error_identity(issue) for issue in previous["issues"]
                                  if issue["severity"] == "error")
        latest_errors = Counter(error_identity(issue) for issue in latest["issues"]
                                if issue["severity"] == "error")
        return any(count > previous_errors[identity]
                   for identity, count in latest_errors.items())

    reason = "QA_PASSED" if current["passed"] else "QA_ISSUES_REMAIN"
    if not current["passed"]:
        if max_seconds is not None and elapsed_seconds >= max_seconds:
            reason = "TIME_LIMIT_REACHED"
        elif max_api_calls is not None and api_calls >= max_api_calls:
            reason = "API_LIMIT_REACHED"
        elif history and (introduces_new_error(history[-1], current) or
                          metrics[-1][1] > metrics[-2][1] or
                          metrics[-1][3] > metrics[-2][3] or
                          scores[-1] < scores[-2] or
                          (scores[-1] == scores[-2] and
                           metrics[-1][1:] > metrics[-2][1:])):
            # Error escalation is unsafe even when clearing many warnings
            # increases the weighted score. Also catch regressions at the
            # score floor where the numeric score cannot fall further.
            reason = "QA_REGRESSED"
        else:
            repeated = 1
            for report in reversed(history):
                if signature(report) != signature(current):
                    break
                repeated += 1
            if repeated >= repeat_limit:
                reason = "QA_STALLED"
            elif len(history) >= max_repairs:
                reason = "REPAIR_LIMIT_REACHED"

    decision = "pass" if reason == "QA_PASSED" else "repair" if reason == "QA_ISSUES_REMAIN" else "stop"
    return {"success": True, "decision": decision, "reason_code": reason,
            "repair_attempts": len(history), "score_delta": scores[-1] - scores[-2] if history else None,
            "suggested_action": "consider_checkpoint_rollback" if reason == "QA_REGRESSED" else None,
            "qa": current}


def combine_qa_reports(structural, checks, criteria):
    """Fold freshly run quality checks into the existing gate's score and history scope."""
    if not isinstance(checks, dict) or not isinstance(criteria, dict):
        raise ValueError("QA checks and criteria must be objects")
    if isinstance(structural, dict) and "checks" in structural:
        raise ValueError("structural QA must be a raw timeline report, not a composite QA report")
    # Never launder a forged structural result into a passing composite report.
    evaluate_qa_gate(structural)
    problems = list(structural["issues"])
    for source, report in checks.items():
        if (not isinstance(source, str) or not source.strip() or len(source) > 64 or
                any(ord(ch) < 32 or ord(ch) == 127 for ch in source)):
            raise ValueError("QA check source must be a printable nonempty name")
        if not isinstance(report, dict) or report.get("success") is not True or \
                not isinstance(report.get("passed"), bool) or not isinstance(report.get("issues"), list) or \
                len(report["issues"]) > 1000:
            raise ValueError(f"{source} did not return a complete QA report")
        _require_json_safe_issues(report["issues"], f"{source} QA")
        for issue in report["issues"]:
            if (not isinstance(issue, dict) or not _valid_qa_code(issue.get("code")) or
                    issue.get("severity") not in ("error", "warning") or
                    ("item_ids" in issue and (not isinstance(issue["item_ids"], list) or
                     any(not _valid_item_id(identity) for identity in issue["item_ids"])))):
                raise ValueError(f"{source} returned an invalid QA issue")
            if issue["code"] == "QA_ISSUES_TRUNCATED" and (
                    issue["severity"] != "error" or
                    type(issue.get("omitted_count")) is not int or
                    issue["omitted_count"] <= 0):
                raise ValueError(f"{source} returned an invalid QA truncation marker")
            if "frame_range" in issue:
                span = issue["frame_range"]
                if (not isinstance(span, list) or len(span) != 2 or
                        type(span[0]) is not int or type(span[1]) is not int or
                        not 0 <= span[0] < span[1] <= MAX_FRAME):
                    raise ValueError(f"{source} returned an invalid QA frame_range")
            normalized = {**issue, "source": source}
            if source == "visual":
                has_start, has_end = "start_frame" in issue, "end_frame" in issue
                if has_start != has_end:
                    raise ValueError("visual QA start_frame and end_frame must be supplied together")
                if has_start:
                    start, end = issue["start_frame"], issue["end_frame"]
                    if (type(start) is not int or type(end) is not int or
                            not 0 <= start < end <= MAX_FRAME or
                            ("frame_range" in issue and issue["frame_range"] != [start, end])):
                        raise ValueError("visual QA returned an invalid or conflicting frame interval")
                    normalized["frame_range"] = [start, end]
            problems.append(normalized)
        has_errors = any(i["severity"] == "error" for i in report["issues"])
        if report["passed"] == has_errors:
            raise ValueError(f"{source} passed status disagrees with its error issues")
        if "valid" in report and report["valid"] is not report["passed"]:
            raise ValueError(f"{source} valid status disagrees with passed")
        if "issue_count" in report and (type(report["issue_count"]) is not int or
                                        report["issue_count"] != len(report["issues"])):
            raise ValueError(f"{source} issue_count disagrees with issues")
        if "problems" in report and report["problems"] != report["issues"]:
            raise ValueError(f"{source} problems disagrees with issues")
        if "summary" in report:
            summary = report["summary"]
            error_count = sum(i["severity"] == "error" for i in report["issues"])
            warning_count = len(report["issues"]) - error_count
            if (not isinstance(summary, dict) or
                    type(summary.get("errors")) is not int or
                    type(summary.get("warnings")) is not int or
                    summary["errors"] != error_count or summary["warnings"] != warning_count):
                raise ValueError(f"{source} summary disagrees with issue severities")
    if set(checks) != set(criteria):
        raise ValueError("QA check sources must match their criteria entries")
    problems = _bounded_qa_issues(problems)
    errors = sum(i["severity"] == "error" for i in problems)
    warnings = sum(i["severity"] == "warning" for i in problems)
    try:
        criteria_json = json.dumps({"structural": structural["criteria_hash"], "checks": criteria},
                                   sort_keys=True, ensure_ascii=False, separators=(",", ":"),
                                   allow_nan=False)
        criteria_bytes = criteria_json.encode("utf-8")
    except (TypeError, ValueError, OverflowError, UnicodeError) as exc:
        raise ValueError("QA criteria must be JSON-safe, finite and UTF-8") from exc
    digest = hashlib.sha256(criteria_bytes).hexdigest()
    return {**structural, "passed": errors == 0, "valid": errors == 0,
            "score": max(0, 100 - errors * 20 - warnings * 5),
            "issue_count": len(problems), "summary": {"errors": errors, "warnings": warnings},
            "issues": problems, "problems": problems, "criteria_hash": digest,
            "checks": checks, "scope": "Timeline structure and requested visual/audio/export checks."}
