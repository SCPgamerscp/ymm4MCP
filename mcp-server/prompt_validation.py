"""Validate MCP production-prompt arguments before reading skill content."""

import ntpath

MAX_THEME_CHARS = 16_384
MAX_VIDEO_PATH_CHARS = 4_096
MAX_DURATION_SECONDS = 86_400


def _windows_media_file(path: str) -> bool:
    """Accept local drive files or complete UNC share files, never devices."""
    drive, tail = ntpath.splitdrive(path)
    if not ntpath.isabs(path) or not tail or not ntpath.basename(path):
        return False
    # Windows treats device names and alternate data streams specially, even
    # when they appear in otherwise absolute file paths.
    reserved = {"CON", "PRN", "AUX", "NUL"}
    reserved.update(f"COM{i}" for i in range(1, 10))
    reserved.update(f"LPT{i}" for i in range(1, 10))
    for part in (segment for segment in tail.replace("/", "\\").split("\\") if segment):
        if (part in {".", ".."} or part.endswith((" ", ".")) or
                any(char in '<>:"|?*' for char in part) or
                part.split(".", 1)[0].rstrip(" ").upper() in reserved):
            return False
    if len(drive) == 2:
        return drive[1] == ":" and drive[0].isascii() and drive[0].isalpha()
    if drive.startswith(ntpath.sep * 2):
        parts = drive.split(ntpath.sep)
        return (len(parts) == 4 and parts[2] not in {"", ".", "?"}
                and bool(parts[3]))
    return False


def validate_prompt_arguments(arguments: object) -> dict:
    args = {} if arguments is None else arguments
    if not isinstance(args, dict):
        raise ValueError("Prompt arguments must be an object")
    if set(args) - {"theme", "video_path", "duration_seconds"}:
        raise ValueError("Unknown prompt argument")

    theme = args.get("theme")
    if (not isinstance(theme, str) or not theme.strip() or
            len(theme) > MAX_THEME_CHARS or
            any((ord(c) < 32 and c not in "\t\n\r") or 127 <= ord(c) <= 159 for c in theme)):
        raise ValueError("theme は必須で、16,384文字以内の有効なテキストで指定してください")

    if "video_path" in args:
        path = args["video_path"]
        if (not isinstance(path, str) or not path.strip() or
                len(path) > MAX_VIDEO_PATH_CHARS or
                any(ord(c) < 32 or ord(c) == 127 for c in path)):
            raise ValueError("video_path は制御文字を含まない空でない文字列で指定してください")
        if not _windows_media_file(path):
            raise ValueError("video_path はWindowsの絶対パスで指定してください")

    if "duration_seconds" in args:
        raw = args["duration_seconds"]
        try:
            duration = int(raw)
        except (ValueError, TypeError, OverflowError) as exc:
            raise ValueError("duration_seconds は1〜86,400の整数秒で指定してください") from exc
        if (isinstance(raw, bool) or not 1 <= duration <= MAX_DURATION_SECONDS or
                str(duration) != str(raw)):
            raise ValueError("duration_seconds は1〜86,400の整数秒で指定してください")
    return args
