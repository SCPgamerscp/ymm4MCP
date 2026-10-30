"""Read-only, whole-timeline verification of a ripple delete (#61).

This does NOT execute edits or prove atomicity. The caller supplies a trusted
pre-edit /items snapshot and this module compares it with a fresh post-edit one.
"""
from ripple_planning import MAX_ITEMS, _frame, _item_id, plan_ripple_delete

MAX_ISSUES = 100
PRESERVED_FIELDS = ("type", "text", "source_path", "character")


def _inventory(snapshot):
    """Validate completeness, identity and geometry before comparing anything."""
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
            frame = _frame(item.get("frame"), "frame")
            length = _frame(item.get("length"), "length", 1)
            _frame(frame + length, "end", 1)
            _frame(item.get("layer"), "layer")
        except ValueError:
            return None
        by_id[identity] = item
    return by_id


def verify_ripple_result(before, after, *, start_frame, end_frame,
                         layers=None, expected_revisions=None):
    """Check exact post-edit item IDs and positions against a fresh preflight.

    The before snapshot must be captured before any edit, not reconstructed
    from the result. Successful verification is observational only: it cannot
    prove an atomic host transaction or saved/reopened project integrity.
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
            if actual[field] != expected:
                issue("RIPPLE_FIELD_MISMATCH", identity, field=field,
                      expected=expected, actual=actual[field])
        for field in PRESERVED_FIELDS:
            if field in item and (field not in actual or actual[field] != item[field]):
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
            "scope": "Snapshot geometry/identity/content comparison only; "
                     "does not prove atomicity, source trimming, save/reopen or visual QA"}
