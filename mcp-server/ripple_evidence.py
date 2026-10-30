"""Short-lived, server-captured evidence for read-only ripple verification (#61).

The MCP client cannot replace the pre-edit inventory when verifying by plan_id.
This is NOT a transaction, edit executor, or proof of YMM4 project identity.
"""
from collections import OrderedDict
from dataclasses import dataclass
import hashlib
import json
import secrets
import time

from ripple_verification import _inventory, verify_ripple_result


@dataclass(frozen=True)
class RippleEvidence:
    created: float
    before_json: bytes
    parameters_json: bytes
    sha256: str


class RippleEvidenceStore:
    TTL_SECONDS = 600
    MAX_RECORDS = 8
    MAX_BYTES = 2_000_000

    def __init__(self, *, clock=None):
        self._clock = clock or time.monotonic
        self._records = OrderedDict()

    def _prune(self):
        now = self._clock()
        for key, record in list(self._records.items()):
            if now - record.created >= self.TTL_SECONDS:
                del self._records[key]

    def capture(self, before, *, start_frame, end_frame, layers,
                expected_revisions):
        """Capture a validated *host-returned* snapshot after successful preflight.

        Call without awaiting between host preflight and capture. This stores
        immutable JSON bytes, never a reference to a mutable caller dictionary.
        """
        if _inventory(before) is None:
            raise ValueError("complete host inventory with count/items is required")
        parameters = dict(start_frame=start_frame, end_frame=end_frame,
                          layers=layers, expected_revisions=expected_revisions)
        try:
            before_json = json.dumps(before, sort_keys=True, ensure_ascii=False,
                                     separators=(",", ":"), allow_nan=False).encode("utf-8")
            parameters_json = json.dumps(parameters, sort_keys=True, ensure_ascii=False,
                                         separators=(",", ":"), allow_nan=False).encode("utf-8")
        except (ValueError, TypeError, OverflowError, UnicodeError, RecursionError) as exc:
            raise ValueError("host inventory is not safe JSON evidence") from exc
        if len(before_json) > self.MAX_BYTES or len(parameters_json) > self.MAX_BYTES:
            raise ValueError("host inventory or plan parameters exceed evidence limit")
        self._prune()
        while len(self._records) >= self.MAX_RECORDS:
            self._records.popitem(last=False)
        token = secrets.token_urlsafe(24)
        digest = hashlib.sha256(before_json).hexdigest()
        self._records[token] = RippleEvidence(self._clock(), before_json,
                                               parameters_json, digest)
        return {"plan_id": token, "before_sha256": digest,
                "evidence_source": "server_captured_host_inventory",
                "evidence_expires_in_seconds": self.TTL_SECONDS}

    def lookup(self, plan_id):
        self._prune()
        if not isinstance(plan_id, str) or len(plan_id) > 128:
            return None
        return self._records.get(plan_id)

    def verify(self, record, after):
        if self._clock() - record.created >= self.TTL_SECONDS:
            return {"success": False, "passed": False,
                    "error_code": "RIPPLE_EVIDENCE_EXPIRED"}
        before = json.loads(record.before_json)
        params = json.loads(record.parameters_json)
        result = verify_ripple_result(before, after, **params)
        result["evidence_source"] = "server_captured_host_inventory"
        result["before_sha256"] = record.sha256
        result["atomicity_verified"] = False
        return result
