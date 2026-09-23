"""Declarative EditPlan: validate, hash, and diff against a timeline snapshot."""
from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path

from editing import integer, validate_timeline
from jobs import MAX_IDEMPOTENCY, is_absolute_media_path

MAX_SCENES = 200
MAX_ITEMS = 1000
MAX_ID_LEN = 64
MAX_TEXT = 10000
KIND = {
    "video": "video", "audio": "audio", "bgm": "audio", "se": "audio",
    "image": "image", "text": "text", "subtitle": "text",
    "dialogue": "voice", "voice": "voice", "tachie": "tachie", "face": "face",
}
ITEM_FIELDS = {
    "video": {"id", "type", "layer", "frame", "length", "source"},
    "audio": {"id", "type", "layer", "frame", "length", "source"},
    "image": {"id", "type", "layer", "frame", "length", "source"},
    "text": {"id", "type", "layer", "frame", "length", "text"},
    "voice": {"id", "type", "layer", "frame", "length", "text", "character"},
    "tachie": {"id", "type", "layer", "frame", "length", "character"},
    "face": {"id", "type", "layer", "frame", "length", "character"},
}


def _id(value, name):
    if not isinstance(value, str) or not 1 <= len(value) <= MAX_ID_LEN or not value.isascii():
        raise ValueError(f"{name} must be a 1..{MAX_ID_LEN} character ASCII id")
    if any(c.isspace() for c in value) or "/" in value or "\\" in value:
        raise ValueError(f"{name} must not contain whitespace or path separators")
    return value


def _kind(type_name):
    if not isinstance(type_name, str):
        raise ValueError("type is required")
    kind = KIND.get(type_name.strip().lower())
    if kind is None:
        raise ValueError(f"unsupported item type: {type_name}")
    return kind


def host_kind(type_name):
    name = (type_name or "").lower()
    for key in ("voice", "video", "audio", "image", "subtitle", "text", "tachie", "face"):
        if key in name:
            return "text" if key == "subtitle" else key
    return name or "unknown"


def parse_plan(args):
    if not isinstance(args, dict):
        raise ValueError("edit plan must be an object")
    raw = args.get("plan", args)
    if not isinstance(raw, dict):
        raise ValueError("plan must be an object")
    scenes = raw.get("scenes")
    if not isinstance(scenes, list) or not 1 <= len(scenes) <= MAX_SCENES:
        raise ValueError(f"scenes must contain 1..{MAX_SCENES} entries")
    project = raw.get("project", {})
    if not isinstance(project, dict):
        raise ValueError("project must be an object")
    extra_project = set(project) - {"fps", "width", "height"}
    if extra_project:
        raise ValueError("project contains unsupported fields")
    fps = integer(project.get("fps", args.get("fps", 30)), "fps", 1, 240)
    width = project["width"] if "width" in project else None
    height = project["height"] if "height" in project else None
    if width is not None:
        integer(width, "width", 16, 7680)
    if height is not None:
        integer(height, "height", 16, 4320)
    gap = integer(raw.get("gap", args.get("gap", 0)), "gap")
    speed = raw.get("chars_per_sec", args.get("chars_per_sec", 5))
    if isinstance(speed, bool) or not isinstance(speed, (int, float)) or not math.isfinite(speed) or not 0 < speed <= 1000:
        raise ValueError("chars_per_sec must be finite and in (0, 1000]")
    start = integer(raw.get("start_frame", args.get("start_frame", 0)), "start_frame")
    key = args.get("idempotency_key", raw.get("idempotency_key"))
    if key is not None:
        if not isinstance(key, str) or not 1 <= len(key) <= MAX_IDEMPOTENCY or key.strip() != key:
            raise ValueError("idempotency_key must be a 1..128 character string")
    if "dry_run" in args and not isinstance(args["dry_run"], bool):
        raise ValueError("dry_run must be boolean")

    parsed_scenes = []
    seen_scenes = set()
    seen_items = set()
    total_items = 0
    cursor = start
    for s_index, scene in enumerate(scenes):
        if not isinstance(scene, dict):
            raise ValueError(f"scenes[{s_index}] must be an object")
        extra_scene = set(scene) - {"id", "duration_policy", "duration", "items"}
        if extra_scene:
            raise ValueError(f"scenes[{s_index}] contains unsupported fields")
        scene_id = _id(scene.get("id"), f"scenes[{s_index}].id")
        if scene_id in seen_scenes:
            raise ValueError(f"duplicate scene id: {scene_id}")
        seen_scenes.add(scene_id)
        policy = scene.get("duration_policy", "fit_content")
        if policy not in {"fit_content", "explicit"}:
            raise ValueError(f"scenes[{s_index}].duration_policy must be fit_content or explicit")
        duration = scene.get("duration")
        if policy == "explicit":
            duration = integer(duration, f"scenes[{s_index}].duration", 1)
        elif duration is not None:
            raise ValueError(f"scenes[{s_index}].duration is only valid with duration_policy=explicit")
        items = scene.get("items")
        if not isinstance(items, list) or not 1 <= len(items) <= MAX_ITEMS:
            raise ValueError(f"scenes[{s_index}].items must contain 1..{MAX_ITEMS} entries")
        total_items += len(items)
        if total_items > MAX_ITEMS:
            raise ValueError(f"plan may contain at most {MAX_ITEMS} items")
        scene_cursor = cursor
        parsed_items = []
        for i_index, item in enumerate(items):
            parsed = _parse_item(item, f"scenes[{s_index}].items[{i_index}]", fps, speed)
            if parsed["id"] in seen_items:
                raise ValueError(f"duplicate item id: {parsed['id']}")
            seen_items.add(parsed["id"])
            if parsed["frame"] is None:
                parsed["frame"] = scene_cursor
                parsed["frame_source"] = "sequential"
            else:
                parsed["frame_source"] = "explicit"
            end = integer(parsed["frame"] + parsed["length"] + gap, "estimated end frame")
            scene_cursor = max(scene_cursor, end)
            parsed_items.append(parsed)
        scene_end = cursor + duration if policy == "explicit" else scene_cursor
        if policy == "explicit" and scene_cursor - cursor > duration:
            raise ValueError(f"scenes[{s_index}] content exceeds explicit duration")
        parsed_scenes.append({
            "id": scene_id, "duration_policy": policy, "duration": duration,
            "start_frame": cursor, "end_frame": scene_end, "items": parsed_items,
        })
        cursor = scene_end
    plan = {
        "project": {k: v for k, v in {"fps": fps, "width": width, "height": height}.items() if v is not None},
        "gap": gap, "chars_per_sec": speed, "start_frame": start, "scenes": parsed_scenes,
        "total_frames": cursor, "item_count": total_items,
    }
    if key is not None:
        plan["idempotency_key"] = key
    return plan


def _parse_item(item, prefix, fps, speed):
    if not isinstance(item, dict):
        raise ValueError(f"{prefix} must be an object")
    kind = _kind(item.get("type"))
    extra = set(item) - ITEM_FIELDS[kind]
    if extra:
        raise ValueError(f"{prefix} contains unsupported fields: {sorted(extra)}")
    parsed = {
        "id": _id(item.get("id"), f"{prefix}.id"),
        "type": item["type"].strip().lower(),
        "kind": kind,
        "layer": integer(item["layer"], f"{prefix}.layer") if "layer" in item else 0,
        "frame": integer(item["frame"], f"{prefix}.frame") if "frame" in item else None,
        "length": None,
        "length_source": "estimated",
    }
    if kind in {"video", "audio", "image"}:
        source = item.get("source")
        if not isinstance(source, str) or not source.strip():
            raise ValueError(f"{prefix}.source is required")
        source = source.strip()
        if not is_absolute_media_path(source):
            raise ValueError(f"{prefix}.source must be an absolute file path")
        parsed["source"] = source
    if kind in {"text", "voice"}:
        text = item.get("text")
        if not isinstance(text, str) or not text.strip() or len(text) > MAX_TEXT:
            raise ValueError(f"{prefix}.text is required and must be 1..{MAX_TEXT} characters")
        parsed["text"] = text
    if kind in {"voice", "tachie", "face"}:
        character = item.get("character")
        if not isinstance(character, str) or not character.strip():
            raise ValueError(f"{prefix}.character is required")
        parsed["character"] = character.strip()
    if "length" in item:
        parsed["length"] = integer(item["length"], f"{prefix}.length", 1)
        parsed["length_source"] = "explicit"
    elif kind == "image":
        raise ValueError(f"{prefix}.length is required for image items")
    elif kind == "voice":
        seconds = max(1.0, len(parsed["text"]) / speed)
        if not math.isfinite(seconds) or seconds > 2_147_483_647 / fps:
            raise ValueError(f"{prefix} estimated duration exceeds the frame range")
        parsed["length"] = max(1, math.ceil(seconds * fps))
        parsed["length_source"] = "estimated"
    else:
        parsed["length"] = 1
        parsed["length_source"] = "deferred" if kind in {"video", "audio"} else "default"
    return parsed


def plan_hash(plan):
    payload = {
        "project": plan["project"], "gap": plan["gap"], "chars_per_sec": plan["chars_per_sec"],
        "start_frame": plan["start_frame"],
        "scenes": [
            {
                "id": scene["id"], "duration_policy": scene["duration_policy"],
                "duration": scene.get("duration"),
                "items": [
                    {k: item[k] for k in ("id", "type", "kind", "layer", "frame", "length",
                                          "source", "text", "character") if k in item}
                    for item in scene["items"]
                ],
            }
            for scene in plan["scenes"]
        ],
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def flatten_items(plan):
    items = []
    for scene in plan["scenes"]:
        for item in scene["items"]:
            items.append({**item, "scene_id": scene["id"]})
    return items


def add_payload(item):
    payload = {"frame": item["frame"], "layer": item["layer"]}
    if item["kind"] == "voice":
        pass
    elif item.get("length") and item.get("length_source") != "deferred":
        payload["length"] = item["length"]
    if "source" in item:
        payload["path"] = item["source"]
    if "text" in item:
        payload["text"] = item["text"]
    if "character" in item:
        payload["character"] = item["character"]
    return payload


def _fingerprint(kind, layer, text=None, character=None, source=None):
    if kind == "voice":
        return ("voice", layer, character, text)
    if kind in {"video", "audio", "image"}:
        return (kind, layer, source or text)
    if kind == "text":
        return ("text", layer, text)
    if kind in {"tachie", "face"}:
        return (kind, layer, character or text)
    return (kind, layer, text)


def current_fingerprint(item):
    kind = host_kind(item.get("type"))
    return _fingerprint(kind, item.get("layer"), text=item.get("text"),
                        character=item.get("character"), source=item.get("path") or item.get("source"))


def desired_fingerprint(item):
    return _fingerprint(item["kind"], item["layer"], text=item.get("text"),
                        character=item.get("character"), source=item.get("source"))


def bindings_map(bindings):
    mapped = {}
    records = bindings
    if isinstance(bindings, dict):
        records = bindings.get("items", [])
    if not isinstance(records, list):
        return mapped
    for entry in records:
        if isinstance(entry, dict) and isinstance(entry.get("id"), str) and isinstance(entry.get("item_id"), str):
            mapped[entry["id"]] = entry
    return mapped


def pick_binding(bindings, key):
    records = []
    if isinstance(bindings, dict):
        if "plan_hash" in bindings or "idempotency_key" in bindings:
            records = [bindings]
        elif isinstance(bindings.get("bindings"), list):
            records = bindings["bindings"]
    elif isinstance(bindings, list):
        records = bindings
    if not key:
        return None
    for record in records:
        if isinstance(record, dict) and record.get("idempotency_key") == key:
            return record
    return None


def replay_record(plan, current_items, record):
    if not isinstance(record, dict):
        return None
    if record.get("plan_hash") != plan_hash(plan):
        return None
    current_ids = {item.get("item_id") for item in current_items if isinstance(item, dict)}
    mapped = bindings_map(record)
    desired = flatten_items(plan)
    if len(mapped) != len(desired):
        return None
    kept = []
    for item in desired:
        bound = mapped.get(item["id"])
        if bound is None or bound.get("item_id") not in current_ids:
            return None
        kept.append({**item, "op": "keep", "item_id": bound["item_id"],
                     "revision": bound.get("revision"), "reason": "idempotent"})
    return {
        "success": True, "replayed": True, "dry_run": False, "added": 0, "kept": len(kept),
        "ops": kept, "warnings": [], "plan_hash": record.get("plan_hash"),
        "idempotency_key": record.get("idempotency_key") or plan.get("idempotency_key"),
        "total_frames": plan["total_frames"], "item_count": plan["item_count"],
        "note": "Same idempotency_key and plan were already applied; nothing was added.",
    }


def diff_plan(plan, current_items=None, bindings=None, character_names=None):
    if current_items is None:
        current_items = []
    if not isinstance(current_items, list):
        raise ValueError("items must be an array")
    names = list(character_names) if character_names else None
    mapped = bindings_map(bindings)
    current_by_id = {item.get("item_id"): item for item in current_items
                     if isinstance(item, dict) and isinstance(item.get("item_id"), str)}
    claimed = set()
    ops = []
    warnings = []
    desired = flatten_items(plan)
    for item in desired:
        if item["kind"] in {"video", "audio", "image"}:
            if not Path(item["source"]).is_file():
                warnings.append({"code": "SOURCE_UNVERIFIED", "severity": "warning", "id": item["id"],
                                 "message": "MCP側からソースファイルを確認できません。YMM4と同じマシン上の絶対パスか確認してください"})
        if item["kind"] in {"voice", "tachie", "face"} and names is not None and names.count(item["character"]) != 1:
            warnings.append({"code": "CHARACTER_UNKNOWN", "severity": "error", "id": item["id"],
                             "character": item["character"],
                             "message": "キャラ名は一覧から一意の完全一致名を指定してください"})
        bound = mapped.get(item["id"])
        if bound and bound.get("item_id") in current_by_id and bound["item_id"] not in claimed:
            current = current_by_id[bound["item_id"]]
            claimed.add(bound["item_id"])
            ops.append({**item, "op": "keep", "item_id": bound["item_id"],
                        "revision": current.get("revision"), "reason": "binding"})
            continue
        match = None
        wanted = desired_fingerprint(item)
        for current in current_items:
            if not isinstance(current, dict) or current.get("item_id") in claimed:
                continue
            if current_fingerprint(current) == wanted:
                if match is not None:
                    match = None
                    break
                match = current
        if match is not None:
            claimed.add(match["item_id"])
            ops.append({**item, "op": "keep", "item_id": match["item_id"],
                        "revision": match.get("revision"), "reason": "content"})
            continue
        ops.append({**item, "op": "add", "payload": add_payload(item), "reason": "missing"})

    for current in current_items:
        if isinstance(current, dict) and current.get("item_id") not in claimed:
            warnings.append({"code": "EXTRA_ITEM", "severity": "warning",
                             "item_id": current.get("item_id"),
                             "frame_range": [current.get("frame"), current.get("endFrame") or current.get("end_frame")],
                             "message": "計画に無い既存アイテムは削除しません"})

    qa_items = [{"item_id": item["id"], "frame": item["frame"], "layer": item["layer"],
                 "length": item["length"]} for item in desired]
    qa = validate_timeline(qa_items, include_gaps=False)
    for issue in qa.get("issues", []):
        if issue.get("code") == "OVERLAP":
            warnings.append({**issue, "severity": issue.get("severity", "error")})

    error_count = sum(w.get("severity") == "error" for w in warnings)
    return {
        "success": True, "replayed": False, "dry_run": True,
        "added": sum(op["op"] == "add" for op in ops),
        "kept": sum(op["op"] == "keep" for op in ops),
        "ops": ops, "warnings": warnings, "passed": error_count == 0,
        "plan_hash": plan_hash(plan), "idempotency_key": plan.get("idempotency_key"),
        "total_frames": plan["total_frames"], "item_count": plan["item_count"],
        "scenes": [{"id": s["id"], "start_frame": s["start_frame"], "end_frame": s["end_frame"],
                    "item_count": len(s["items"])} for s in plan["scenes"]],
        "note": "No timeline edits performed. Voice lengths are estimates until apply.",
    }


def to_edit_state(items, bindings=None):
    if not isinstance(items, list):
        raise ValueError("items must be an array")
    mapped = []
    for item in items:
        if not isinstance(item, dict):
            continue
        mapped.append({
            "id": item.get("item_id"),
            "item_id": item.get("item_id"),
            "revision": item.get("revision"),
            "identity_persistent": item.get("identity_persistent"),
            "type": host_kind(item.get("type")),
            "host_type": item.get("type"),
            "layer": item.get("layer"),
            "frame": item.get("frame"),
            "length": item.get("length"),
            "text": item.get("text"),
            "character": item.get("character"),
            "source": item.get("path") or item.get("source"),
        })
    return {
        "success": True,
        "scenes": [{"id": "timeline", "duration_policy": "fit_content", "items": mapped}],
        "item_count": len(mapped),
        "bindings": bindings_map(bindings),
        "note": "Host timeline has no scene labels; items are returned as one scene.",
    }
