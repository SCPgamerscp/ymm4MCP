"""Request validation and output-file checks for export/jobs. No YMM4 required."""
from __future__ import annotations

import ntpath
import re
from pathlib import Path

ALLOWED_FORMATS = {"mp4", "wav", "avi", "mov", "mkv", "webm"}
MAX_IDEMPOTENCY = 128
MAX_TIMEOUT = 7200
DEFAULT_TIMEOUT = 1800
_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_JOB_ID = re.compile(r"^job_[A-Za-z0-9]{8,64}$")


def _valid_windows_file_path(candidate: str) -> bool:
    """Reject Windows device namespaces, incomplete UNC paths and unsafe segments."""
    drive, tail = ntpath.splitdrive(candidate)
    if not ntpath.isabs(candidate) or not tail:
        return False
    if len(drive) == 2 and drive[1] == ":":
        if not drive[0].isascii() or not drive[0].isalpha():
            return False
    else:
        # UNC must have a real server and share; device/extended namespaces
        # (\\.\ and \\?\) must never reach file-writing host endpoints.
        parts = re.split(r"[\\/]", drive[2:])
        if not drive.startswith(("\\\\", "//")) or len(parts) != 2 or (
                parts[0] in {"", ".", "?"} or not parts[1]):
            return False
    if tail[0] not in "\\/":
        return False
    segments = re.split(r"[\\/]", tail[1:])
    if not segments or any(not part or part in {".", ".."} for part in segments):
        return False
    for part in segments:
        if (part.endswith((" ", ".")) or any(char in '<>:"|?*' for char in part) or
                re.fullmatch(r"(?i)(?:con|prn|aux|nul|conin\$|conout\$|com[1-9]|lpt[1-9])(?:\..*)?", part)):
            return False
    return True


def is_absolute_media_path(path: str) -> bool:
    if not isinstance(path, str) or not path.strip() or any(ord(c) < 32 or ord(c) == 127 for c in path):
        return False
    candidate = path.strip()
    if candidate.startswith(("\\\\", "//")) or _DRIVE.match(candidate):
        return _valid_windows_file_path(candidate)
    return Path(candidate).is_absolute()


def job_id_ok(job_id: str) -> bool:
    return isinstance(job_id, str) and bool(_JOB_ID.fullmatch(job_id))


def validate_export_request(args: dict) -> dict:
    if not isinstance(args, dict):
        raise ValueError("export request must be an object")
    path = args.get("path", args.get("output_path"))
    if not isinstance(path, str) or not path.strip():
        raise ValueError("path は書き出し先の絶対パスで指定してください")
    path = path.strip()
    if not is_absolute_media_path(path):
        raise ValueError("path must be an absolute file path")
    suffix = Path(path.replace("\\", "/")).suffix.lower().lstrip(".")
    fmt = args.get("format", suffix)
    if isinstance(fmt, str):
        fmt = fmt.lower().lstrip(".")
    if fmt not in ALLOWED_FORMATS:
        raise ValueError("format must be one of: " + ", ".join(sorted(ALLOWED_FORMATS)))
    if suffix and suffix != fmt:
        raise ValueError("path extension must match format")
    overwrite = args.get("overwrite", False)
    if not isinstance(overwrite, bool):
        raise ValueError("overwrite must be boolean")
    timeout = args.get("timeout_seconds", DEFAULT_TIMEOUT)
    if isinstance(timeout, bool) or not isinstance(timeout, int) or not 1 <= timeout <= MAX_TIMEOUT:
        raise ValueError(f"timeout_seconds must be an integer in 1..{MAX_TIMEOUT}")
    key = args.get("idempotency_key")
    if key is not None:
        if (not isinstance(key, str) or not 1 <= len(key) <= MAX_IDEMPOTENCY
                or key.strip() != key or any(ord(c) < 32 or ord(c) == 127 for c in key)):
            raise ValueError("idempotency_key must be a 1..128 character string")
    return {
        "output_path": path,
        "format": fmt,
        "overwrite": overwrite,
        "timeout_seconds": timeout,
        "idempotency_key": key,
    }


def validate_project_path(args: dict, must_exist: bool = False) -> str:
    del must_exist
    path = args.get("path")
    if not isinstance(path, str) or not path.strip():
        raise ValueError("path はプロジェクトの絶対パスで指定してください")
    path = path.strip()
    if not is_absolute_media_path(path):
        raise ValueError("path must be an absolute file path")
    suffix = Path(path.replace("\\", "/")).suffix.lower()
    if suffix not in {".ymmp"}:
        raise ValueError("project path must end with .ymmp")
    return path


def _ebml_vint(data: bytes, offset: int) -> tuple[int, int] | None:
    """Read a bounded EBML variable-length integer (size field)."""
    if offset >= len(data) or data[offset] == 0:
        return None
    mask = 0x80
    length = 1
    while not data[offset] & mask:
        mask >>= 1
        length += 1
        if length > 8:
            return None
    if offset + length > len(data):
        return None
    value = data[offset] & (mask - 1)
    for byte in data[offset + 1:offset + length]:
        value = (value << 8) | byte
    return value, offset + length


def _ebml_element_id(data: bytes, offset: int) -> tuple[bytes, int] | None:
    """Read a bounded EBML element ID without consuming payload bytes."""
    if offset >= len(data) or data[offset] == 0:
        return None
    mask, length = 0x80, 1
    while not data[offset] & mask:
        mask >>= 1
        length += 1
        if length > 4:
            return None
    if offset + length > len(data):
        return None
    return data[offset:offset + length], offset + length


def _ebml_doctype(header: bytes) -> str | None:
    """Walk EBML header children, never matching markers inside Void payloads."""
    if header[:4] != b"\x1a\x45\xdf\xa3":
        return None
    root = _ebml_vint(header, 4)
    if root is None:
        return None
    root_size, begin = root
    if root_size == (1 << (7 * (begin - 4))) - 1:
        return None  # Unknown-size headers cannot be bounded safely.
    end = begin + root_size
    if end > len(header):
        return None
    cursor = begin
    while cursor < end:
        element = _ebml_element_id(header, cursor)
        if element is None:
            return None
        element_id, size_offset = element
        length = _ebml_vint(header, size_offset)
        if length is None:
            return None
        size, start = length
        if size == (1 << (7 * (start - size_offset))) - 1:
            return None
        next_element = start + size
        if next_element > end:
            return None
        if element_id == b"\x42\x82":
            if not 1 <= size <= 32:
                return None
            try:
                return header[start:next_element].decode("ascii")
            except UnicodeDecodeError:
                return None
        cursor = next_element
    return None


def _mov_has_metadata_atom(path: Path, file_size: int) -> bool:
    """Skip top-level QuickTime atoms by size, even when mdat precedes moov.

    Reads at most 16 bytes per atom and never loads media payload into memory.
    """
    try:
        with path.open("rb") as stream:
            offset = 0
            for _ in range(128):
                if offset + 8 > file_size:
                    return False
                stream.seek(offset)
                atom = stream.read(16)
                if len(atom) < 8:
                    return False
                size = int.from_bytes(atom[:4], "big")
                minimum = 8
                if size == 1:
                    if len(atom) < 16:
                        return False
                    size = int.from_bytes(atom[8:16], "big")
                    minimum = 16
                elif size == 0:
                    size = file_size - offset
                if size < minimum or offset + size > file_size:
                    return False
                if atom[4:8] == b"moov":
                    return True
                offset += size
                if offset == file_size:
                    return False
    except OSError:
        return False
    return False


def verify_media_file(path: str, expected_format: str | None = None) -> dict:
    if not isinstance(path, str) or not path.strip():
        return {"ok": False, "error_code": "EXPORT_FILE_MISSING", "error": "出力パスが空です"}
    target = Path(path)
    missing = {"ok": False, "error_code": "EXPORT_FILE_MISSING",
               "error": "出力ファイルがありません", "output_path": path}
    try:
        if not target.is_file():
            return missing
        with target.open("rb") as stream:
            size = stream.seek(0, 2)
            stream.seek(0)
            # Bounded read: never load the potentially multi-gigabyte export.
            header = stream.read(4096)
    except FileNotFoundError:
        return missing
    except OSError:
        return {"ok": False, "error_code": "EXPORT_VERIFY_FAILED",
                "error": "出力ファイルを読み取れません", "output_path": path}
    suffix = target.suffix.lstrip(".").lower()
    fmt = expected_format if expected_format is not None else suffix
    if not isinstance(fmt, str):
        fmt = ""
    fmt = fmt.lower().lstrip(".")

    def fail(message: str) -> dict:
        return {"ok": False, "error_code": "EXPORT_VERIFY_FAILED", "error": message,
                "bytes": size, "output_path": path, "format": fmt}

    if fmt not in ALLOWED_FORMATS or (suffix and suffix != fmt):
        return fail("出力ファイル形式と拡張子が一致しません")
    if size < 32:
        return {"ok": False, "error_code": "EXPORT_FILE_EMPTY", "error": "出力ファイルが小さすぎます",
                "bytes": size, "output_path": path, "format": fmt}
    result = {"ok": True, "format": fmt, "bytes": size, "output_path": path,
              "verified": True, "verification_scope": "container_header_only",
              "streams_verified": False}

    if fmt in {"mp4", "mov"}:
        if fmt == "mov" and header[4:8] != b"ftyp":
            if not _mov_has_metadata_atom(target, size):
                return fail("QuickTimeのmoovメタデータが見つからないか、atomサイズが不正です")
            return result
        box_type = header[4:8]
        if box_type != b"ftyp" and not (fmt == "mov" and box_type == b"moov"):
            return fail("ISO BMFF/QuickTimeの先頭ボックスが不正です")
        box_size = int.from_bytes(header[:4], "big")
        if box_size == 1:
            box_size = int.from_bytes(header[8:16], "big")
            minimum_size, brand_offset = (24, 16) if box_type == b"ftyp" else (16, 0)
        else:
            minimum_size, brand_offset = (16, 8) if box_type == b"ftyp" else (8, 0)
        if not minimum_size <= box_size <= size:
            return fail("ISO BMFF/QuickTimeの先頭ボックスサイズが不正です")
        if box_type == b"ftyp":
            result["brand"] = header[brand_offset:brand_offset + 4].decode("latin1", "replace")
        if fmt == "mp4":
            result["has_video"] = True  # Legacy compatibility; stream presence is not verified.
        return result
    if fmt in {"wav", "avi"}:
        expected_tag = b"WAVE" if fmt == "wav" else b"AVI "
        if header[:4] != b"RIFF" or header[8:12] != expected_tag:
            return fail("RIFFヘッダが不正です")
        riff_size = int.from_bytes(header[4:8], "little")
        if riff_size < 4 or riff_size + 8 > size:
            return fail("RIFFチャンクサイズが不正です")
        if fmt == "wav":
            result["has_audio"] = True
        return result
    if fmt in {"mkv", "webm"}:
        expected_type = "matroska" if fmt == "mkv" else "webm"
        if _ebml_doctype(header) != expected_type:
            return fail("EBMLヘッダのDocTypeが出力形式と一致しません")
        return result
    return fail("未対応の出力形式です")
