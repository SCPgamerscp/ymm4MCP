# Post-edit item verification (Issue #61)

The MCP action `ymm4_interact` with `action="verify_items"` reads one
`GET /api/items` snapshot and compares specific fields of items selected
by their unique `item_id`. It does not change the YMM4 project.

Example arguments:

```json
{
  "action": "verify_items",
  "expected": [
    {"item_id": "native:EXAMPLE_ID", "revision": "EXAMPLE_REVISION", "frame": 120, "layer": 2, "length": 90},
    {"item_id": "native:ANOTHER_ID", "text": "こんにちは"}
  ]
}
```

Replace the example IDs and revisions with values obtained from the *current*
`get_info/items` response. A missing item, mismatched field, unreadable
field, duplicate item ID, incomplete snapshot, or invalid expectation cannot
produce `passed=true`. Results contain `success` (the snapshot was usable),
`passed`/`verified` (all requested fields matched), `matched_count`,
`checked_count`, and machine-readable `issues` with the affected field.
An invalid request is rejected before the host is contacted.

This is **not** a transaction lock: a manual UI edit can race the read. A
successful result is only evidence about that one host snapshot. It does not
prove persistence after saving/reopening, correct keyframe interpolation,
preview rendering, or full rollback. For Issue #140, continue to run the
separate YMM4 4.56.1.1 live keyframe acceptance test. For Issue #61, keep
full-transaction rollback and rendered-video QA as separate requirements.
