# Ripple delete preflight (Issue #61, partial implementation)

`ymm4_interact(action="plan_ripple_delete", start_frame=100, end_frame=150,
layers=[3], expected_revisions={"native:clip1": "REV1", "native:clip2": "REV2"})`
is a **read-only** preflight for removing the half-open timeline range
`[100,150)` from selected layers. Obtain exact item IDs and revisions from
`get_info/items` immediately beforehand. If `layers` is omitted, every
layer is in scope. This does **not** edit or save the YMM4 project.

The result lists proposed `delete` operations for items fully inside the
range, and `move` operations for items starting at or after the end.
An item crossing a cut boundary returns `RIPPLE_SPLIT_REQUIRED` rather than
silently shortening the source. Missing, stale or ambiguous item IDs/revisions
return `RIPPLE_CONFLICT` or `RIPPLE_ID_AMBIGUOUS`. Malformed snapshots are
rejected. Plans are sorted by layer/frame/ID to make review reproducible.

**Safety limit:** `success=true` means the plan passed static preflight,
not that the edit is complete or even executable. The response deliberately
includes `executable=false`. Do **not** replay the returned operations as
separate `edit_item` requests: a failure halfway through could destroy data.
The C# YMM4 host still needs a single atomic ripple-delete transaction,
source-aware split/trim for boundary-crossing items, verified full rollback,
revision rechecks immediately before mutation, and live Windows/YMM4
acceptance tests. Issue #61 must remain open until those are implemented and
verified.

Local pure-Python tests:
`python -m unittest -v test_ripple_planning` (from `mcp-server`).

## Post-edit whole-timeline verification (read-only)

Before any *separately implemented atomic host edit*, capture a complete
`ymm4_interact(action="get_info", sub_action="items")` response, including
`count`, `items`, stable `item_id` and `revision` fields. Retain that
**original** snapshot. After the edit, call:

```json
{
  "action": "verify_ripple_delete",
  "before_snapshot": {"success": true, "count": 2, "items": [
    {"item_id": "native:cut", "frame": 10, "length": 10, "layer": 1, "revision": "r1"},
    {"item_id": "native:after", "frame": 20, "length": 10, "layer": 1, "revision": "r2"}
  ]},
  "start_frame": 10,
  "end_frame": 20,
  "layers": [1],
  "expected_revisions": {"native:cut": "r1", "native:after": "r2"}
}
```

The MCP server fetches the current `/api/items` snapshot **without writing**
and compares the complete ID set, planned deletions, shifts, unaffected item
positions/layers/lengths, and **every observed non-positional field**.
This now includes nested keyframes, effects, audio levels, source offsets,
opacity, and any future fields exposed by the host, rather than only a
hardcoded set of text/source fields. Missing/added fields and type changes
are reported as `RIPPLE_CONTENT_CHANGED`. The only excluded comparison
fields are `frame` (checked against the plan), `end_frame` **or the
actual YMM4 host spelling `endFrame`** (if present, must equal
`frame + length` and retain field presence), and `revision` (which may
advance after a legitimate edit).

Inventories with inconsistent `count`, duplicate IDs, impossible
`end_frame`, non-finite JSON values or invalid Unicode are rejected.
The response contains `passed` and machine-readable `issues`;
`executed=false` and `atomicity_verified=false` are deliberate.
**A pass only checks fields actually exposed in both snapshots.** Host data
that is not serialized in `/api/items` cannot be verified here.

**Do not confuse this with an implementation of ripple deletion.** The
`before_snapshot` is caller-supplied and is not cryptographically authenticated.
A pass cannot establish that the operation was atomic, that source trim offsets
were correct, that the saved project can be reopened, or that the rendered video
is valid. These still require host transaction/rollback and live YMM4 acceptance.

Local regression tests:
`python -m unittest -v test_ripple_verification test_ripple_integrity_extended`.

## Server-owned pre-edit evidence (recommended; still read-only)

The original `verify_ripple_delete(before_snapshot=...)` accepts a caller-provided
pre-edit inventory. That input can be accidentally stale or replaced by a
client; a passing result therefore does **not** establish where the original
inventory came from.

To bind a verification to an inventory actually read from the YMM4 host by
the MCP server, request a short-lived plan handle:

```json
{
  "action": "plan_ripple_delete",
  "start_frame": 10,
  "end_frame": 20,
  "layers": [1],
  "expected_revisions": {"native:cut": "REV1", "native:after": "REV2"},
  "capture_evidence": true
}
```

When preflight succeeds and the host supplies a **complete** inventory with
matching `count` and `items`, the response includes `plan_id`,
`before_sha256`, `evidence_source=server_captured_host_inventory`, and
`evidence_expires_in_seconds=600`. The server stores an immutable copy of the
host response and original cut parameters, not a reference to a client object.

After a separately implemented edit, verify with **only**:

```json
{"action": "verify_ripple_delete", "plan_id": "<plan_id from preflight>"}
```

The server fetches fresh `/items` and verifies against its own pre-edit
snapshot. Supplying `before_snapshot`, cut bounds, layers, or revisions with
`plan_id` is rejected rather than silently overriding the captured plan.

Evidence is kept **in memory only** (up to 8 recent plans, 2 MB per snapshot,
10-minute TTL). An unknown, evicted, expired, or restart-lost `plan_id`
returns `RIPPLE_EVIDENCE_UNAVAILABLE`; if it expires during a host request,
`RIPPLE_EVIDENCE_EXPIRED` is returned. Capture fails closed when the host
inventory is incomplete or too large. Do not retry the edit just because
verification evidence expired.

**This is evidence provenance, not an atomic edit implementation.** It cannot
prove project identity after switching projects, prevent another process from
changing the timeline between reads, restore deleted items, or verify
properties omitted by the host. Actual ripple execution, full rollback,
save/reopen and visual/audio acceptance are still missing. Keep #61 open.

Tests: `python -m unittest -v test_ripple_evidence
test_ripple_evidence_dispatch` from `mcp-server`.
