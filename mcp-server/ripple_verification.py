"""Read-only, fail-closed whole-timeline verification of ripple deletion (#61).

This does NOT execute edits or prove atomicity. The caller supplies a trusted
pre-edit /items snapshot and this module compares it with a fresh post-edit one.
"""
import json
from ripple_planning import MAX_ITEMS, _frame, _item_id, plan_ripple_delete

MAX_ISSUES = 100
# A move is allowed to change its position, derived end, and optimistic revision.
# All other observed properties (including nested effects/animations/volume)
# must remain identical. Missing/added fields are NOT silently ignored.
MUTABLE_FIELDS = frozenset(("frame", "end_frame", "revision"))


def _canonical(value):
    """Type-sensitive, deterministic JSON representation; reject unsafe evidence."""
    return json.dumps(value, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":"), allow_nan=False).encode("utf-8")


def _inventory(snapshot):
    """Validate completeness, identity, geometry and JSON safety."""
    if not isinstance(snapshot, dict) or snapshot.get("success") is False or (
            "error" in snapshot or "error_code" in snapshot):
        return None
    items = snapshot.get("items")
    count = snapshot.get("count")
    if (not isinstance(items, list) or len(items) > MAX_ITEMS
            or type(count) is not int or count != len(items)):
        return None
    by_id = {}
    for item in items:
        if not isinstance(item, dict) or not _item_id(item.get("item_id")):
            return None
        identity = item["item_id"]
        if identity in by_id:
            return None
        try:
            _canonical(item)
            frame = _frame(item.get("frame"), "frame")
            length = _frame(item.get("length"), "length", 1)
            _frame(frame + length, "end", 1)
            _frame(item.get("layer"), "layer")
            if "end_frame" in item and (
                    type(item["end_frame"]) is not int or
                    item["end_frame"] != frame + length):
                return None
        except (ValueError, TypeError, OverflowError, UnicodeError):
            return None
        by_id[identity] = item
    return by_id


def verify_ripple_result(before, after, *, start_frame, end_frame,
                         layers=None, expected_revisions=None):
    """Verify exact item IDs, geometry and ALL other observed properties.

    Caller-supplied before evidence is not authenticated. This cannot prove
    host atomicity, persisted project integrity, or visual/audio correctness.
    """
    old = _inventory(before)
    if old is None:
        return {"success": False, "passed": False,
                "error_code": "RIPPLE_BEFORE_SNAPSHOT_INVALID"}
    plan = plan_ripple_delete(before, start_frame=start_frame, end_frame=end_frame,
                              layers=layers, expected_revisions=expected_revisions)
    if not plan["success"]:
        return {"success": False, "passed": False,
                "error_code": "RIPPLE_PREFLIGHT_FAILED", "preflight": plan}
    new = _inventory(after)
    if new is None:
        return {"success": False, "passed": False,
                "error_code": "RIPPLE_AFTER_SNAPSHOT_INVALID"}

    operations = {op["item_id"]: op for op in plan["operations"]}
    issues = []
    omitted = 0

    def issue(code, identity, **details):
        nonlocal omitted
        if len(issues) < MAX_ISSUES:
            issues.append({"code": code, "item_id": identity, **details})
        else:
            omitted += 1

    for identity, item in old.items():
        operation = operations.get(identity)
        if operation and operation["action"] == "delete":
            if identity in new:
                issue("RIPPLE_ITEM_NOT_DELETED", identity)
            continue
        actual = new.get(identity)
        if actual is None:
            issue("RIPPLE_UNEXPECTED_DELETION", identity)
            continue
        expected_frame = operation["to_frame"] if operation else item["frame"]
        for field, expected in (("frame", expected_frame),
                                ("layer", item["layer"]),
                                ("length", item["length"])):
            if type(actual[field]) is not type(expected) or actual[field] != expected:
                issue("RIPPLE_FIELD_MISMATCH", identity, field=field,
                      expected=expected, actual=actual[field])
        # A derived end frame must follow the move, never remain at its old
        # position. _inventory already validates consistency within each item.
        if ("end_frame" in item) != ("end_frame" in actual):
            issue("RIPPLE_CONTENT_CHANGED", identity, field="end_frame")
        # Compare all observed host data, not just text and source_path.
        # This catches unintended changes to effects, keyframes, opacity,
        # volume, source offsets, and any future fields added by YMM4.
        for field in sorted((set(item) | set(actual)) - MUTABLE_FIELDS):
            if field not in item or field not in actual or (
                    _canonical(item[field]) != _canonical(actual[field])):
                issue("RIPPLE_CONTENT_CHANGED", identity, field=field)

    for identity in new:
        if identity not in old:
            issue("RIPPLE_UNEXPECTED_ADDITION", identity)
    if omitted:
        issues.append({"code": "RIPPLE_ISSUES_TRUNCATED", "omitted_count": omitted})
    return {"success": True, "passed": not issues, "verified": not issues,
            "checked_before_count": len(old), "checked_after_count": len(new),
            "planned_delete_count": plan["delete_count"],
            "planned_move_count": plan["move_count"], "issues": issues,
            "executed": False, "atomicity_verified": False,
            "scope": "Read-only complete snapshot geometry/identity/all-observed-field comparison; "
                     "does not prove atomicity, source trimming, save/reopen or visual QA"}
