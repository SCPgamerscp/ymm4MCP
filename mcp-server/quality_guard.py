"""Read-only consistency checks shared by quality gates and derived edits."""

import editplan


async def verify_snapshot(get, items):
    expected = editplan.snapshot_hash(items)
    try:
        current = await get("/items")
        if (not isinstance(current, dict) or current.get("success") is False or
                "error" in current or not isinstance(current.get("items"), list)):
            return {"success": False, "error_code": "ITEMS_UNAVAILABLE", "details": current}
        actual = editplan.snapshot_hash(current["items"])
    except Exception as exc:
        return {"success": False, "error_code": "SNAPSHOT_VERIFY_UNKNOWN", "error": str(exc)}
    if actual != expected:
        return {"success": False, "error_code": "SNAPSHOT_CONFLICT",
                "expected_snapshot_hash": expected, "actual_snapshot_hash": actual}
    return None
