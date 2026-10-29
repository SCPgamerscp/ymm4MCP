"""Safe, opt-in overlap organizer for YMM4 timeline items.

Preview is the default. Repeated dialogue is intentional and is never deleted.
Run with --apply only after reviewing the proposed moves.
"""
import argparse
import asyncio
import json
from pathlib import Path
import sys
from urllib.parse import urlsplit

import httpx

MAX_FRAME = 2_147_483_647


def snapshot(items):
    """Validate the host snapshot and return stable identity/revision tuples."""
    if not isinstance(items, list) or len(items) > 10000:
        raise ValueError("items must be an array of at most 10000 entries")
    result = {}
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("each item must be an object")
        item_id, revision = item.get("item_id"), item.get("revision")
        if not isinstance(item_id, str) or not item_id.strip() or item_id in result:
            raise ValueError("items require unique nonempty item_id values")
        if not isinstance(revision, str) or not revision:
            raise ValueError("each item requires a revision for safe editing")
        values = []
        for key, minimum in (("layer", 0), ("frame", 0), ("length", 1)):
            value = item.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
                raise ValueError(f"{item_id}.{key} must be a valid integer")
            values.append(value)
        layer, frame, length = values
        if frame + length > MAX_FRAME:
            raise ValueError(f"{item_id} exceeds the supported frame range")
        result[item_id] = (layer, frame, length, revision)
    return result


def plan_moves(items):
    """Move later overlapping items right, preserving each item's duration."""
    current = snapshot(items)
    by_layer = {}
    for item_id, (layer, frame, length, revision) in current.items():
        by_layer.setdefault(layer, []).append((frame, item_id, length, revision))
    moves = []
    for layer in sorted(by_layer):
        cursor = 0
        for frame, item_id, length, revision in sorted(by_layer[layer]):
            target = max(frame, cursor)
            if target + length > MAX_FRAME:
                raise ValueError("overlap repair would exceed the supported frame range")
            if target != frame:
                moves.append({"item_id": item_id, "layer": layer, "old_frame": frame,
                              "new_frame": target, "expected_revision": revision})
            cursor = target + length
    if len(moves) > 500:
        raise ValueError("more than 500 moves; split the project into smaller batches")
    return moves


async def get_items(client, base, headers):
    response = await client.get(base + "/items", headers=headers)
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict) or payload.get("success") is False or "error" in payload:
        raise ValueError("YMM4 rejected the timeline snapshot")
    items = payload.get("items")
    snapshot(items)
    return items


async def apply_moves(client, base, headers, initial_items, moves):
    """Fail closed on stale state or unverifiable edits; never delete items."""
    expected = snapshot(initial_items)
    if snapshot(await get_items(client, base, headers)) != expected:
        raise RuntimeError("timeline changed since planning; no edits were started")
    completed = 0
    for move in moves:
        item_id = move["item_id"]
        layer, frame, length, revision = expected[item_id]
        if (frame != move["old_frame"] or layer != move["layer"]
                or revision != move["expected_revision"]):
            raise RuntimeError(f"stale plan for {item_id}; stopped after {completed} moves")
        # A different editor or user may change the timeline between writes.
        if snapshot(await get_items(client, base, headers)) != expected:
            raise RuntimeError(f"timeline changed; stopped after {completed} moves")
        response = await client.post(base + "/items/prop", headers=headers, json={
            "item_id": item_id, "expected_revision": revision,
            "prop": "Frame", "value": str(move["new_frame"])
        })
        response.raise_for_status()
        result = response.json()
        if (not isinstance(result, dict) or result.get("success") is not True
                or result.get("verified") is not True
                or result.get("item_id") != item_id
                or not isinstance(result.get("revision"), str)):
            raise RuntimeError(f"move not verified for {item_id}; inspect YMM4 before retrying")
        expected[item_id] = (layer, move["new_frame"], length, result["revision"])
        if snapshot(await get_items(client, base, headers)) != expected:
            raise RuntimeError(f"post-edit snapshot mismatch for {item_id}; stop and inspect YMM4")
        completed += 1
    return completed


async def run(apply=False):
    # Reuse the MCP client's authenticated loopback-only connection descriptor.
    sys.path.insert(0, str(Path(__file__).resolve().parent / "mcp-server"))
    from ymm4_connection import connection_settings

    base, headers = connection_settings()
    if urlsplit(base).port is None:
        raise ValueError("YMM4_API_BASE must contain an explicit local port")
    async with httpx.AsyncClient(timeout=10.0, trust_env=False) as client:
        items = await get_items(client, base, headers)
        moves = plan_moves(items)
        print(json.dumps({"dry_run": not apply, "move_count": len(moves),
                          "moves": moves}, ensure_ascii=False, indent=2))
        if not apply or not moves:
            return 0
        completed = await apply_moves(client, base, headers, items, moves)
        print(f"Verified {completed} item moves. No items were deleted.")
        return 0


def main():
    parser = argparse.ArgumentParser(description="Preview safe overlap repairs; no text deduplication.")
    parser.add_argument("--apply", action="store_true",
                        help="explicitly apply the previewed shifts with revision checks")
    args = parser.parse_args()
    try:
        return asyncio.run(run(args.apply))
    except (ValueError, RuntimeError, httpx.HTTPError, OSError) as exc:
        print(f"Timeline organizer stopped: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
