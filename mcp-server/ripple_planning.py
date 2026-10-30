"""Read-only, deterministic preflight for future atomic timeline ripple deletion.

No YMM4 write endpoint is called by this module. The operation list is an
audit artifact, NOT a command queue: sequential replay is not atomic.
"""
MAX_FRAME = 2_147_483_647
MAX_ITEMS = 10_000


def _frame(value, label, minimum=0):
    if type(value) is not int or not minimum <= value <= MAX_FRAME:
        raise ValueError(f"{label} must be an integer in {minimum}..{MAX_FRAME}")
    return value


def _item_id(value):
    return (isinstance(value, str) and 1 <= len(value) <= 512
            and value == value.strip()
            and not any(ord(c) < 32 or ord(c) == 127 or
                        0xD800 <= ord(c) <= 0xDFFF for c in value))


def _revision(value):
    return (isinstance(value, str) and 1 <= len(value) <= 256
            and value == value.strip()
            and all(33 <= ord(c) <= 126 for c in value))


def _conflict(code, message, **details):
    return {"success": False, "dry_run": True, "executable": False,
            "error_code": code, "message": message, **details}


def plan_ripple_delete(snapshot, *, start_frame, end_frame,
                       layers=None, expected_revisions=None):
    """Preflight [start_frame, end_frame) without mutating the host.

    Fully enclosed items are deleted; later items are shifted left. Items
    crossing a cut boundary need source-aware split/trim and are rejected.
    Each affected item needs a caller-supplied matching optimistic revision.
    """
    start = _frame(start_frame, "start_frame")
    end = _frame(end_frame, "end_frame", 1)
    if start >= end:
        raise ValueError("start_frame must precede end_frame")
    if layers is not None:
        if not isinstance(layers, list) or not 1 <= len(layers) <= 128:
            raise ValueError("layers must contain 1..128 unique layer integers")
        selected = set(_frame(layer, "layers entry") for layer in layers)
        if len(selected) != len(layers):
            raise ValueError("layers must not contain duplicates")
    else:
        selected = None
    if expected_revisions is None:
        expected_revisions = {}
    if not isinstance(expected_revisions, dict) or len(expected_revisions) > MAX_ITEMS:
        raise ValueError("expected_revisions must be a bounded item_id-to-revision map")
    if any(not _item_id(k) or not _revision(v) for k, v in expected_revisions.items()):
        raise ValueError("expected_revisions must contain valid item IDs and revisions")
    if (not isinstance(snapshot, dict) or snapshot.get("success") is False
            or "error" in snapshot or "error_code" in snapshot
            or not isinstance(snapshot.get("items"), list)
            or len(snapshot["items"]) > MAX_ITEMS):
        return _conflict("RIPPLE_SNAPSHOT_UNAVAILABLE", "Complete timeline inventory is required")

    seen, operations, affected = set(), [], []
    boundary_conflicts, revision_conflicts = [], []
    delta = end - start
    for index, item in enumerate(snapshot["items"]):
        if not isinstance(item, dict):
            return _conflict("RIPPLE_SNAPSHOT_INVALID", "Malformed timeline item", index=index)
        identity = item.get("item_id")
        if not _item_id(identity) or identity in seen:
            return _conflict("RIPPLE_ID_AMBIGUOUS", "Missing, invalid or duplicate item ID", index=index)
        seen.add(identity)
        try:
            frame = _frame(item.get("frame"), "item frame")
            length = _frame(item.get("length"), "item length", 1)
            layer = _frame(item.get("layer"), "item layer")
            _frame(frame + length, "item end", 1)
        except ValueError:
            return _conflict("RIPPLE_SNAPSHOT_INVALID", "Invalid timeline geometry", index=index)
        if selected is not None and layer not in selected:
            continue
        tail = frame + length
        if tail <= start:
            continue
        if frame < end and tail > start and not (frame >= start and tail <= end):
            boundary_conflicts.append({"item_id": identity, "frame": frame,
                                       "end_frame": tail, "layer": layer})
            continue
        if frame >= end:
            operation = {"action": "move", "item_id": identity, "layer": layer,
                         "from_frame": frame, "to_frame": frame - delta,
                         "length": length}
        elif frame >= start and tail <= end:
            operation = {"action": "delete", "item_id": identity, "layer": layer,
                         "frame": frame, "length": length}
        else:
            continue
        affected.append(identity)
        revision = item.get("revision")
        expected = expected_revisions.get(identity)
        if not _revision(revision) or expected != revision:
            revision_conflicts.append({"item_id": identity,
                                       "reason": "missing_revision" if not _revision(revision)
                                       else "expected_revision_mismatch"})
        else:
            operation["expected_revision"] = revision
        operations.append(operation)

    if boundary_conflicts:
        return _conflict("RIPPLE_SPLIT_REQUIRED",
                         "Boundary-crossing items need source-aware split/trim before ripple deletion",
                         conflicts=boundary_conflicts)
    if revision_conflicts:
        return _conflict("RIPPLE_CONFLICT",
                         "Affected items require matching optimistic revisions",
                         conflicts=revision_conflicts, affected_item_ids=affected)
    operations.sort(key=lambda op: (op["layer"], op.get("frame", op.get("from_frame")),
                                    op["item_id"]))
    return {"success": True, "dry_run": True, "executable": False,
            "start_frame": start, "end_frame": end,
            "layers": sorted(selected) if selected is not None else None,
            "operations": operations,
            "delete_count": sum(op["action"] == "delete" for op in operations),
            "move_count": sum(op["action"] == "move" for op in operations),
            "requires": "atomic_host_transaction_with_verified_full_rollback",
            "note": "Read-only preflight only; do not replay these operations as individual writes."}
