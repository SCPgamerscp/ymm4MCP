"""Plan bounded Volume keyframes around the current voice spans."""

from editing import finite_number, integer


def plan(bgm: dict, voices: list[dict], base_volume: float, *, ratio: float = .3,
         attack_frames: int = 5, release_frames: int = 10) -> list[dict]:
    if not isinstance(bgm, dict):
        raise ValueError("bgm must be an object")
    ratio = finite_number(ratio, "duck_ratio")
    if not 0 < ratio < 1:
        raise ValueError("duck_ratio must be in (0, 1)")
    attack_frames = integer(attack_frames, "attack_frames", 1, 300)
    release_frames = integer(release_frames, "release_frames", 1, 300)
    base_volume = finite_number(base_volume, "base_volume")
    if base_volume <= 0:
        raise ValueError("base_volume must be positive")
    start = integer(bgm.get("frame"), "bgm.frame")
    length = integer(bgm.get("length"), "bgm.length", 1)
    end = integer(start + length, "bgm.end")
    if not isinstance(voices, list) or len(voices) > 500:
        raise ValueError("at most 500 voice items are supported")
    spans = []
    for index, voice in enumerate(voices):
        if not isinstance(voice, dict):
            raise ValueError(f"voices[{index}] must be an object")
        frame = integer(voice.get("frame"), "voice.frame")
        duration = integer(voice.get("length"), "voice.length", 1)
        tail = integer(frame + duration, "voice.end")
        if frame < end and tail > start:
            spans.append((max(frame, start), min(tail, end)))
    merged = []
    for left, right in sorted(spans):
        if merged and left - merged[-1][1] <= attack_frames + release_frames:
            merged[-1] = (merged[-1][0], max(right, merged[-1][1]))
        else:
            merged.append((left, right))
    points = {}
    low = base_volume * ratio
    for left, right in merged:
        attack = max(start, left - attack_frames)
        if attack < left:
            points[attack - start] = base_volume
        points[left - start] = low
        if right < end:
            points[right - start] = low
            # Do not jump back to full volume at the final audible frame when
            # there is not enough room for the requested release transition.
            if right + release_frames < end:
                points[right + release_frames - start] = base_volume
    if len(points) > 2000:
        raise ValueError("ducking plan exceeds 2000 keyframes")
    return [{"at": at, "value": value} for at, value in sorted(points.items())]


def verify_keyframes(report: dict, revision: str, points: list[dict], base_volume: float) -> bool:
    """Verify the complete curve, including the unchanged original start key."""
    if (not isinstance(report, dict) or report.get("success") is not True or
            "error" in report or "error_code" in report or
            report.get("revision") != revision or not isinstance(report.get("animations"), list)):
        return False
    volume = [a for a in report["animations"] if isinstance(a, dict) and a.get("prop") == "Volume"]
    found = volume[0].get("keyframes") if len(volume) == 1 else None
    # A non-finite expected value (especially NaN) must never verify as equal:
    # abs(actual - NaN) >= tolerance is False for *every* actual value.
    if not isinstance(points, list) or isinstance(base_volume, bool) or not isinstance(base_volume, (int, float)):
        return False
    try:
        base_volume = finite_number(base_volume, "base_volume")
        expected = {}
        for point in points:
            if not isinstance(point, dict):
                return False
            at = integer(point.get("at"), "keyframe.at")
            raw_value = point.get("value")
            if at in expected or isinstance(raw_value, bool) or not isinstance(raw_value, (int, float)):
                return False
            expected[at] = finite_number(raw_value, "keyframe.value")
        expected.setdefault(0, base_volume)
    except (ValueError, OverflowError):
        return False
    if not isinstance(found, list) or len(found) != len(expected):
        return False
    seen = set()
    for key in found:
        if not isinstance(key, dict) or not isinstance(key.get("value"), (int, float)):
            return False
        try:
            at = integer(key.get("at"), "keyframe.at")
            value = finite_number(key["value"], "keyframe.value")
        except ValueError:
            return False
        if at in seen or at not in expected or abs(value - expected[at]) >= 1e-6:
            return False
        seen.add(at)
    return True
