"""Read-only consistency checks shared by quality gates and derived edits."""

import editplan
from editing import integer


def _validate_observable_fields(items):
    """Do not certify a snapshot whose observable timeline metadata is invalid.

    snapshot_hash deliberately fingerprints fields without coercion, but two
    equally malformed snapshots would otherwise compare equal and pass QA.
    Keep fields optional for older YMM4 hosts; validate them when supplied.
    """
    if not isinstance(items, list):
        return  # snapshot_hash reports malformed container shapes.
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            continue  # snapshot_hash reports malformed rows.
        for field in ("frame", "layer", "length"):
            if field in item:
                integer(item[field], f"items[{index}].{field}",
                        1 if field == "length" else 0)
        for field in ("revision", "type"):
            if field in item:
                value = item[field]
                if (not isinstance(value, str) or not value.strip()
                        or len(value) > 256
                        or any(ord(char) < 32 or ord(char) == 127 for char in value)):
                    raise ValueError(f"items[{index}].{field} must be a nonempty safe string")


async def verify_snapshot(get, items):
    # The caller's reference snapshot can be malformed too. Fail closed
    # instead of raising before we have a chance to return a QA failure.
    try:
        _validate_observable_fields(items)
        expected = editplan.snapshot_hash(items)
    except Exception as exc:
        return {"success": False, "error_code": "SNAPSHOT_VERIFY_UNKNOWN", "error": str(exc)}
    try:
        current = await get("/items")
        if (not isinstance(current, dict) or
                ("success" in current and current["success"] is not True) or
                "error" in current or "error_code" in current or
                not isinstance(current.get("items"), list)):
            return {"success": False, "error_code": "ITEMS_UNAVAILABLE", "details": current}
        _validate_observable_fields(current["items"])
        actual = editplan.snapshot_hash(current["items"])
    except Exception as exc:
        return {"success": False, "error_code": "SNAPSHOT_VERIFY_UNKNOWN", "error": str(exc)}
    if actual != expected:
        return {"success": False, "error_code": "SNAPSHOT_CONFLICT",
                "expected_snapshot_hash": expected, "actual_snapshot_hash": actual}
    return None
