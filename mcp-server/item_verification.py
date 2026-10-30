"""Read-only, identity-based post-edit verification of the live YMM4 timeline.

Unlike structural QA's expected-item matcher, this compares individual fields
of a uniquely identified item and reports which field failed. It never guesses
an item from frame/layer coordinates and never mutates the host.
"""
from collections import Counter

MAX_EXPECTED = 1000
MAX_SNAPSHOT = 10000
FIELDS = frozenset({"revision", "frame", "layer", "length", "type", "text"})


def _valid_id(value):
    return (isinstance(value, str) and 1 <= len(value) <= 512 and
            value == value.strip() and
            all(32 <= ord(c) < 127 or ord(c) >= 160 for c in value) and
            not any(0xD800 <= ord(c) <= 0xDFFF for c in value))


def validate_expected_items(expected):
    """Reject ambiguous/empty expectations before making a host request."""
    if not isinstance(expected, list) or not 1 <= len(expected) <= MAX_EXPECTED:
        raise ValueError("verify_items expected must contain 1..1000 entries")
    seen = set()
    for index, entry in enumerate(expected):
        if not isinstance(entry, dict) or not set(entry).issubset(FIELDS | {"item_id"}):
            raise ValueError(f"expected[{index}] has unsupported fields")
        item_id = entry.get("item_id")
        if not _valid_id(item_id):
            raise ValueError(f"expected[{index}].item_id must be a valid stable ID")
        if item_id in seen:
            raise ValueError(f"expected[{index}].item_id is repeated")
        seen.add(item_id)
        if len(entry) == 1:
            raise ValueError(f"expected[{index}] must specify at least one field to verify")
        for field, value in entry.items():
            if field in ("frame", "layer", "length"):
                if isinstance(value, bool) or not isinstance(value, int) or not (
                    (1 if field == "length" else 0) <= value <= 2_147_483_647):
                    raise ValueError(f"expected[{index}].{field} must be a valid frame/layer/length")
            elif field in ("revision", "type"):
                if not isinstance(value, str) or not 1 <= len(value) <= 512 or not value.strip():
                    raise ValueError(f"expected[{index}].{field} must be a nonempty string")
            elif field == "text":
                if not isinstance(value, str) or len(value) > 100_000:
                    raise ValueError(f"expected[{index}].text must be a string of at most 100000 characters")
            if isinstance(value, str):
                try:
                    value.encode("utf-8")
                except UnicodeError as exc:
                    raise ValueError(f"expected[{index}].{field} is not valid UTF-8") from exc
    return expected


def verify_items(snapshot, expected):
    """Compare one complete /api/items snapshot with explicit stable-ID expectations.

    success means the snapshot was valid and comparison ran; passed means every
    requested field matched. A missing/ambiguous item is never treated as a pass.
    """
    validate_expected_items(expected)
    if (not isinstance(snapshot, dict) or snapshot.get("success") is False or
            "error" in snapshot or "error_code" in snapshot):
        return {"success": False, "passed": False, "error_code": "ITEMS_UNAVAILABLE"}
    items = snapshot.get("items")
    count = snapshot.get("count")
    if (not isinstance(items, list) or len(items) > MAX_SNAPSHOT or
            isinstance(count, bool) or not isinstance(count, int) or count != len(items)):
        return {"success": False, "passed": False, "error_code": "ITEMS_SNAPSHOT_INVALID"}
    ids = []
    for item in items:
        if not isinstance(item, dict) or not _valid_id(item.get("item_id")):
            return {"success": False, "passed": False, "error_code": "ITEMS_SNAPSHOT_INVALID"}
        ids.append(item["item_id"])
    duplicates = [item_id for item_id, n in Counter(ids).items() if n > 1]
    if duplicates:
        return {"success": False, "passed": False, "error_code": "ITEM_ID_AMBIGUOUS",
                "duplicate_item_ids": duplicates[:20]}
    by_id = {item["item_id"]: item for item in items}
    issues = []
    verified = 0
    for wanted in expected:
        item_id = wanted["item_id"]
        actual = by_id.get(item_id)
        if actual is None:
            issues.append({"code": "ITEM_MISSING", "severity": "error", "item_id": item_id})
            continue
        matched = True
        for field, value in wanted.items():
            if field == "item_id":
                continue
            if field not in actual:
                issues.append({"code": "ITEM_FIELD_UNAVAILABLE", "severity": "error",
                               "item_id": item_id, "field": field})
                matched = False
            elif type(actual[field]) is not type(value) or actual[field] != value:
                issues.append({"code": "ITEM_FIELD_MISMATCH", "severity": "error",
                               "item_id": item_id, "field": field,
                               "expected": value, "actual": actual[field]})
                matched = False
        if matched:
            verified += 1
    return {"success": True, "passed": not issues, "verified": not issues,
            "checked_count": len(expected), "matched_count": verified,
            "snapshot_count": count, "issues": issues,
            "scope": "Current /api/items snapshot; not a save/reopen or visual verification"}
