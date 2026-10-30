"""Validate MCP production-prompt arguments before reading skill content."""

import ntpath
import re
import unicodedata

MAX_THEME_CHARS = 16_384
MAX_VIDEO_PATH_CHARS = 4_096
MAX_DURATION_SECONDS = 86_400


# Directional formatting controls can make a path appear to refer to a
# different filename in logs or in a human approval prompt.
_BIDI_CONTROLS = {0x061C, 0x200E, 0x200F, *range(0x202A, 0x202F),
                  *range(0x2066, 0x206A)}
_WINDOWS_DEVICES = {"CON", "PRN", "AUX", "NUL"}
_WINDOWS_DEVICES.update(f"{prefix}{digit}" for prefix in ("COM", "LPT")
                        for digit in "123456789¹²³")


def _invalid_unicode(char: str, *, path: bool) -> bool:
    code = ord(char)
    # Isolated UTF-16 surrogates and Unicode noncharacters are not valid
    # portable JSON/UTF-8 text. Keep ordinary emoji and RTL text valid.
    if (0xD800 <= code <= 0xDFFF or 0xFDD0 <= code <= 0xFDEF
            or (code & 0xFFFE) == 0xFFFE):
        return True
    if path and (unicodedata.category(char) == "Cc" or code in _BIDI_CONTROLS):
        return True
    return False


def _valid_windows_component(part: str) -> bool:
    if (not part or part in {".", ".."} or part.endswith((" ", "."))
            or any(char in '<>:"|?*' or _invalid_unicode(char, path=True)
                   for char in part)):
        return False
    return part.split(".", 1)[0].rstrip(" ").upper() not in _WINDOWS_DEVICES


def _windows_media_file(path: str) -> bool:
    """Accept ordinary drive files and complete UNC share files, never devices."""
    drive, tail = ntpath.splitdrive(path)
    if not ntpath.isabs(path) or not tail or not ntpath.basename(path):
        return False
    # Reject doubled separators and empty components instead of normalizing
    # them to a different file identity. Both drive and UNC tails start with
    # exactly one path separator after splitdrive().
    normalized_tail = tail.replace("/", "\\")
    if not normalized_tail.startswith("\\") or normalized_tail.startswith("\\\\"):
        return False
    if not all(_valid_windows_component(part)
               for part in normalized_tail[1:].split("\\")):
        return False
    if len(drive) == 2:
        return drive[1] == ":" and drive[0].isascii() and drive[0].isalpha()
    if drive.startswith(("\\\\", "//")):
        # ntpath.splitdrive puts the server and share in *drive*, not tail.
        # Validate both, otherwise \\server\..\file and malformed UNC host
        # names bypass all per-component checks above.
        parts = drive.replace("/", "\\").split("\\")
        return (len(parts) == 4 and parts[:2] == ["", ""]
                and all(_valid_windows_component(part) for part in parts[2:]))
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
            any((ord(c) < 32 and c not in "\t\n\r") or 127 <= ord(c) <= 159
                or _invalid_unicode(c, path=False) for c in theme)):
        raise ValueError("theme は必須で、16,384文字以内の有効なテキストで指定してください")

    if "video_path" in args:
        path = args["video_path"]
        if (not isinstance(path, str) or not path.strip() or
                len(path) > MAX_VIDEO_PATH_CHARS or
                any(_invalid_unicode(c, path=True) for c in path)):
            raise ValueError("video_path は制御文字を含まない空でない文字列で指定してください")
        if not _windows_media_file(path):
            raise ValueError("video_path はWindowsの絶対パスで指定してください")

    if "duration_seconds" in args:
        raw = args["duration_seconds"]
        # Avoid invoking arbitrary __int__/__str__ hooks and converting huge
        # untrusted decimal strings. JSON integers and canonical ASCII strings
        # are the only supported representations.
        if type(raw) is int:
            duration = raw
        elif type(raw) is str and re.fullmatch(r"[1-9][0-9]{0,4}", raw):
            duration = int(raw)
        else:
            raise ValueError("duration_seconds は1〜86,400の整数秒で指定してください")
        if not 1 <= duration <= MAX_DURATION_SECONDS:
            raise ValueError("duration_seconds は1〜86,400の整数秒で指定してください")
    return args
