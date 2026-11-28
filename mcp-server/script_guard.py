"""Optional optimistic snapshot check for script placement."""
import re

import editplan
import quality_guard
from editing import plan_script


async def read_snapshot(get, expected=None):
    """Return the item list only after checking the host inventory and any expected hash."""
    try:
        snapshot = await get("/items")
    except Exception as exc:
        return None, {"success": False, "error_code": "SNAPSHOT_VERIFY_UNKNOWN", "error": str(exc)}
    failure = quality_guard.timeline_inventory_error(snapshot)
    if failure is not None:
        return None, failure
    if expected is not None:
        try:
            actual = editplan.snapshot_hash(snapshot["items"])
        except Exception as exc:
            return None, {"success": False, "error_code": "SNAPSHOT_VERIFY_UNKNOWN", "error": str(exc)}
        if expected != actual:
            return None, {"success": False, "error_code": "SNAPSHOT_CONFLICT",
                          "expected_snapshot_hash": expected, "actual_snapshot_hash": actual,
                          "added": 0}
    return snapshot["items"], None


async def apply(args, get, add_script):
    expected = args.get("expected_snapshot_hash")
    if expected is None:
        return await add_script(args)
    if not isinstance(expected, str) or re.fullmatch(r"[0-9a-f]{64}", expected) is None:
        raise ValueError("expected_snapshot_hash must be a lowercase SHA-256 hex digest")
    # Planning requires dry_run=True, but snapshot preflight must also validate
    # real writes without changing the arguments sent to the host.
    plan_script({**args, "dry_run": True})  # Side-effect-free input validation.
    _, failure = await read_snapshot(get, expected)
    if failure is not None:
        return failure
    # Recheck inside add_script after character/FPS/append preflight, immediately
    # before its first voice write. The first check alone can become stale there.
    result = await add_script(args, expected_snapshot=expected)
    # A malformed or contradictory result cannot certify a successful write.
    if (not isinstance(result, dict) or
            (result.get("success") is not True and result.get("success") is not False) or
            (result.get("success") is True
             and ("error" in result or "error_code" in result))):
        return {"success": False, "error_code": "SCRIPT_RESULT_INVALID",
                "error": "Script host returned a malformed or contradictory result"}
    # A failed dry run must not be decorated with a trusted snapshot marker.
    if result["success"] is True and args.get("dry_run"):
        return {**result, "snapshot_hash": expected}
    return result
