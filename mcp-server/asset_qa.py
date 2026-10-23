"""Turn read-only host asset checks into an explicit final QA criterion."""


def report(availability: dict) -> dict:
    if not isinstance(availability, dict) or availability.get("success") is not True:
        raise ValueError("asset availability check failed")
    missing = availability.get("missing")
    unknown = availability.get("unknown_source_item_ids")
    unchecked = availability.get("unchecked")
    if not all(isinstance(value, list) for value in (missing, unknown, unchecked)):
        raise ValueError("asset availability report is incomplete")
    issues = []
    for source in missing:
        issues.append({"code": "ASSET_MISSING", "severity": "error", "path": source["path"],
                       "item_ids": source["item_ids"]})
    if unknown:
        issues.append({"code": "ASSET_SOURCE_UNKNOWN", "severity": "error", "item_ids": unknown})
    for source in unchecked:
        issues.append({"code": "ASSET_CHECK_UNAVAILABLE", "severity": "error", "path": source["path"],
                       "item_ids": source["item_ids"]})
    if availability.get("complete") is not True and not issues:
        raise ValueError("incomplete asset check has no explanation")
    return {"success": True, "passed": not issues, "issues": issues,
            "checked_count": availability.get("checked_count")}
