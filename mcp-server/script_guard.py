"""Optional optimistic snapshot check for script placement."""
import re

import editplan
from editing import plan_script


async def apply(args, get, add_script):
    expected = args.get("expected_snapshot_hash")
    if expected is None:
        return await add_script(args)
    if not isinstance(expected, str) or re.fullmatch(r"[0-9a-f]{64}", expected) is None:
        raise ValueError("expected_snapshot_hash must be a lowercase SHA-256 hex digest")
    plan_script(args)  # Reject malformed scripts before calling the host.
    snapshot = await get("/items")
    if not isinstance(snapshot, dict) or snapshot.get("success") is False or \
            not isinstance(snapshot.get("items"), list):
        return {"success": False, "error_code": "ITEMS_UNAVAILABLE", "details": snapshot}
    actual = editplan.snapshot_hash(snapshot["items"])
    if expected != actual:
        return {"success": False, "error_code": "SNAPSHOT_CONFLICT",
                "expected_snapshot_hash": expected, "actual_snapshot_hash": actual}
    result = await add_script(args)
    if isinstance(result, dict) and args.get("dry_run"):
        return {**result, "snapshot_hash": actual}
    return result
