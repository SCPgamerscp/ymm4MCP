"""Request validation and output-file checks for export/jobs. No YMM4 required."""
from __future__ import annotations

import re
from pathlib import Path

ALLOWED_FORMATS = {"mp4", "wav", "avi", "mov", "mkv", "webm"}
MAX_IDEMPOTENCY = 128
MAX_TIMEOUT = 7200
DEFAULT_TIMEOUT = 1800
_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_JOB_ID = re.compile(r"^job_[A-Za-z0-9]{8,64}$")


def is_absolute_media_path(path: str) -> bool:
    if not isinstance(path, str) or not path.strip() or "\x00" in path or "\n" in path:
        return False
    candidate = path.strip()
    if candidate.startswith("\\\\") or _DRIVE.match(candidate):
        return True
    return Path(candidate).is_absolute()


def job_id_ok(job_id: str) -> bool:
    return isinstance(job_id, str) and bool(_JOB_ID.match(job_id))


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
        if not isinstance(key, str) or not 1 <= len(key) <= MAX_IDEMPOTENCY or key.strip() != key:
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


def verify_media_file(path: str, expected_format: str | None = None) -> dict:
    if not isinstance(path, str) or not path.strip():
        return {"ok": False, "error_code": "EXPORT_FILE_MISSING", "error": "出力パスが空です"}
    target = Path(path)
    if not target.is_file():
        return {"ok": False, "error_code": "EXPORT_FILE_MISSING", "error": "出力ファイルがありません", "output_path": path}
    size = target.stat().st_size
    fmt = (expected_format or target.suffix.lstrip(".")).lower()
    if size < 32:
        return {"ok": False, "error_code": "EXPORT_FILE_EMPTY", "error": "出力ファイルが小さすぎます",
                "bytes": size, "output_path": path, "format": fmt}
    header = target.read_bytes()[:12]
    if fmt == "mp4":
        if header[4:8] != b"ftyp":
            return {"ok": False, "error_code": "EXPORT_VERIFY_FAILED", "error": "MP4のftypボックスがありません",
                    "bytes": size, "output_path": path, "format": fmt}
        brand = header[8:12].decode("latin1", "replace")
        return {"ok": True, "format": "mp4", "bytes": size, "output_path": path,
                "has_video": True, "brand": brand, "verified": True}
    if fmt == "wav":
        if header[:4] != b"RIFF" or header[8:12] != b"WAVE":
            return {"ok": False, "error_code": "EXPORT_VERIFY_FAILED", "error": "WAVヘッダが不正です",
                    "bytes": size, "output_path": path, "format": fmt}
        return {"ok": True, "format": "wav", "bytes": size, "output_path": path,
                "has_audio": True, "verified": True}
    return {"ok": True, "format": fmt, "bytes": size, "output_path": path, "verified": True}
