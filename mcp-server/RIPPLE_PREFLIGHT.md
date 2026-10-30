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
positions/layers/lengths and selected content fields. It rejects incomplete
inventories, duplicate IDs, stale preflight revisions, missing/unexpected items
and changed content. The response contains `passed` and machine-readable
`issues`; `executed=false` and `atomicity_verified=false` are deliberate.

**Do not confuse this with an implementation of ripple deletion.** The
`before_snapshot` is caller-supplied and is not cryptographically authenticated.
A pass cannot establish that the operation was atomic, that source trim offsets
were correct, that the saved project can be reopened, or that the rendered video
is valid. These still require host transaction/rollback and live YMM4 acceptance.
