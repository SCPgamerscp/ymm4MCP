"""Validate MCP production-prompt arguments before reading skill content."""

import ntpath

MAX_THEME_CHARS = 16_384
MAX_VIDEO_PATH_CHARS = 4_096
MAX_DURATION_SECONDS = 86_400


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
        drive, _ = ntpath.splitdrive(path)
        if not (ntpath.isabs(path) and
                (len(drive) == 2 and drive[0].isascii() and drive[0].isalpha()
                 and drive[1] == ":" or drive.startswith("\\\\"))):
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
