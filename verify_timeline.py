"""Read-only timeline QA. Repeated dialogue is informational, not an error."""
import asyncio
import json
from pathlib import Path
import sys
from urllib.parse import urlsplit

import httpx

MAX_FRAME = 2_147_483_647


def inspect(items):
    if not isinstance(items, list) or len(items) > 10000:
        raise ValueError("items must be an array of at most 10000 entries")
    by_layer, seen_ids, seen_text, repeats = {}, set(), {}, []
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("each item must be an object")
        item_id = item.get("item_id")
        if not isinstance(item_id, str) or not item_id.strip() or item_id in seen_ids:
            raise ValueError("items require unique nonempty item_id values")
        seen_ids.add(item_id)
        values = []
        for key, minimum in (("layer", 0), ("frame", 0), ("length", 1)):
            value = item.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
                raise ValueError(f"{item_id}.{key} must be an integer >= {minimum}")
            values.append(value)
        layer, frame, length = values
        if frame + length > MAX_FRAME:
            raise ValueError(f"{item_id} exceeds the supported frame range")
        by_layer.setdefault(layer, []).append((frame, frame + length, item_id))
        item_type, text = item.get("type"), item.get("text")
        if item_type in ("VoiceItem", "TextItem") and isinstance(text, str) and text:
            key = (item_type, text)
            if key in seen_text:
                repeats.append({"item_id": item_id, "previous_item_id": seen_text[key],
                                "type": item_type})
            else:
                seen_text[key] = item_id

    issues = []
    for layer, entries in sorted(by_layer.items()):
        active_end, active_id = -1, None
        for start, end, item_id in sorted(entries):
            if start < active_end:
                issues.append({"code": "SAME_LAYER_OVERLAP", "layer": layer,
                               "item_ids": [active_id, item_id],
                               "overlap_start_frame": start,
                               "overlap_end_frame": min(active_end, end)})
            # Keep the furthest end, not merely the previous item: nested
            # intervals otherwise hide later overlaps.
            if end > active_end:
                active_end, active_id = end, item_id
    return {"passed": not issues, "item_count": len(items),
            "overlap_count": len(issues), "issues": issues,
            "repeated_dialogue": repeats,
            "note": "Repeated text is not an error. Same-layer overlap is structural only."}


async def verify():
    sys.path.insert(0, str(Path(__file__).resolve().parent / "mcp-server"))
    from ymm4_connection import connection_settings

    base, headers = connection_settings()
    if urlsplit(base).port is None:
        raise ValueError("YMM4_API_BASE must specify a local port")
    async with httpx.AsyncClient(timeout=10.0, trust_env=False) as client:
        response = await client.get(base + "/items", headers=headers)
        response.raise_for_status()
        payload = response.json()
    if not isinstance(payload, dict) or payload.get("success") is False or "error" in payload:
        raise ValueError("YMM4 rejected the timeline snapshot")
    return inspect(payload.get("items"))


def main():
    try:
        result = asyncio.run(verify())
    except (ValueError, RuntimeError, httpx.HTTPError, OSError) as exc:
        print(json.dumps({"passed": False, "error": str(exc)}, ensure_ascii=False),
              file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
