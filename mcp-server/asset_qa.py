"""Turn read-only host asset checks into an explicit final QA criterion."""

import ntpath

from jobs import is_absolute_media_path


def _valid_item_id(value: object) -> bool:
    return (isinstance(value, str) and bool(value.strip())
            and not any(ord(char) < 32 or ord(char) == 127 for char in value))


def report(availability: dict) -> dict:
    if (not isinstance(availability, dict) or availability.get("success") is not True or
            "error" in availability or "error_code" in availability):
        raise ValueError("asset availability check failed")
    missing = availability.get("missing")
    unknown = availability.get("unknown_source_item_ids")
    unchecked = availability.get("unchecked")
    if not all(isinstance(value, list) for value in (missing, unknown, unchecked)):
        raise ValueError("asset availability report is incomplete")
    complete = availability.get("complete")
    count = availability.get("checked_count")
    if not isinstance(complete, bool) or complete != (not unknown and not unchecked):
        raise ValueError("asset availability completeness is inconsistent")
    if type(count) is not int or count < 0:
        raise ValueError("checked_count must be a non-negative integer")

    # Host paths are grouped by canonical Windows path: a source cannot be
    # both missing and unchecked, and an item cannot have multiple statuses.
    # Reject contradictory snapshots instead of reporting a misleading QA result.
    seen_paths = set()
    seen_item_ids = set()
    for item_id in unknown:
        if not _valid_item_id(item_id) or item_id in seen_item_ids:
            raise ValueError("unknown source item IDs must be unique valid strings")
        seen_item_ids.add(item_id)

    for source in [*missing, *unchecked]:
        if (not isinstance(source, dict) or not isinstance(source.get("path"), str) or
                source["path"].strip() != source["path"] or
                not is_absolute_media_path(source["path"])):
            raise ValueError("asset source entries require a valid path")
        item_ids = source.get("item_ids")
        if not isinstance(item_ids, list) or not item_ids:
            raise ValueError("asset source entries require item IDs")
        canonical_path = ntpath.normcase(ntpath.normpath(source["path"]))
        if canonical_path in seen_paths:
            raise ValueError("asset source path appears in multiple states")
        seen_paths.add(canonical_path)
        for item_id in item_ids:
            if not _valid_item_id(item_id) or item_id in seen_item_ids:
                raise ValueError("asset item IDs must be unique valid strings")
            seen_item_ids.add(item_id)

    if len(missing) > count:
        raise ValueError("missing asset count exceeds checked_count")

    issues = []
    for source in missing:
        issues.append({"code": "ASSET_MISSING", "severity": "error", "path": source["path"],
                       "item_ids": source["item_ids"]})
    if unknown:
        issues.append({"code": "ASSET_SOURCE_UNKNOWN", "severity": "error", "item_ids": unknown})
    for source in unchecked:
        issues.append({"code": "ASSET_CHECK_UNAVAILABLE", "severity": "error", "path": source["path"],
                       "item_ids": source["item_ids"]})
    if not complete and not issues:
        raise ValueError("incomplete asset check has no explanation")
    return {"success": True, "passed": not issues, "issues": issues,
            "checked_count": count}
