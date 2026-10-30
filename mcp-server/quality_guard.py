"""Read-only consistency checks shared by quality gates and derived edits."""

import editplan
from editing import MAX_FRAME, integer

MAX_SNAPSHOT_ITEMS = 100_000
_BIDI_CONTROLS = {0x061C, 0x200E, 0x200F, *range(0x202A, 0x202F),
                  *range(0x2066, 0x206A)}


def _unsafe_metadata_character(char, *, multiline=False):
    code = ord(char)
    return (code in _BIDI_CONTROLS or 0xD800 <= code <= 0xDFFF
            or 0xFDD0 <= code <= 0xFDEF or (code & 0xFFFE) == 0xFFFE
            or (code < 32 and not (multiline and char in '\t\n\r'))
            or 127 <= code <= 159)


def _validate_observable_fields(items):
    """Do not certify a snapshot whose observable timeline metadata is invalid.

    snapshot_hash deliberately fingerprints fields without coercion, but two
    equally malformed snapshots would otherwise compare equal and pass QA.
    Keep fields optional for older YMM4 hosts; validate them when supplied.
    """
    if not isinstance(items, list):
        return  # snapshot_hash reports malformed container shapes.
    if len(items) > MAX_SNAPSHOT_ITEMS:
        raise ValueError(f"snapshot exceeds {MAX_SNAPSHOT_ITEMS} items")
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            continue  # snapshot_hash reports malformed rows.
        for field in ("frame", "layer", "length"):
            if field in item:
                integer(item[field], f"items[{index}].{field}",
                        1 if field == "length" else 0)
        if "endFrame" in item:
            end = integer(item["endFrame"], f"items[{index}].endFrame",
                          0, 2 * MAX_FRAME)
            if ("frame" in item and "length" in item
                    and end != item["frame"] + item["length"]):
                raise ValueError(f"items[{index}].endFrame contradicts frame + length")
        for field in ("revision", "type"):
            if field in item:
                value = item[field]
                if (not isinstance(value, str) or not value.strip()
                        or len(value) > 256
                        or any(_unsafe_metadata_character(char) for char in value)):
                    raise ValueError(f"items[{index}].{field} must be a nonempty safe string")
        # Optional text can contain newlines, but neither JSON objects nor
        # Unicode spoofing controls are trustworthy host timeline metadata.
        for field, limit, multiline in (("text", 65_536, True),
                                        ("source_path", 4_096, False)):
            if field in item and item[field] is not None:
                value = item[field]
                if (type(value) is not str or len(value) > limit
                        or any(_unsafe_metadata_character(c, multiline=multiline)
                               for c in value)):
                    raise ValueError(f"items[{index}].{field} must be safe text")


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
