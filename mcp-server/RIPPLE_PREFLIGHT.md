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
