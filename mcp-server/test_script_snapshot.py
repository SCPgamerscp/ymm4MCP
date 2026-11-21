"""Guard script placement against edits between preview and apply."""
import unittest
from unittest.mock import AsyncMock, call, patch

import editplan
import server


class ScriptSnapshotTests(unittest.IsolatedAsyncioTestCase):
    async def test_change_during_character_preflight_stops_before_first_write(self):
        empty_hash = editplan.snapshot_hash([])
        args = {"action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                "expected_snapshot_hash": empty_hash}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "items": []},
                {"success": True, "characters": [{"name": "A"}]},
                {"success": True, "items": [{"item_id": "other", "frame": 0, "length": 10}]}])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch(args)
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")
        self.assertEqual(result["added"], 0)
        post.assert_not_awaited()

    async def test_conflict_stops_before_voice_synthesis(self):
        args = {"action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                "expected_snapshot_hash": editplan.snapshot_hash([])}
        changed = {"success": True, "items": [{"item_id": "x", "frame": 0, "layer": 0, "length": 10}]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=changed)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch(args)
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")
        post.assert_not_awaited()

    async def test_dry_run_returns_observed_hash(self):
        args = {"action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                "expected_snapshot_hash": editplan.snapshot_hash([]), "dry_run": True}
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "items": []})), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch(args)
        self.assertEqual(result["snapshot_hash"], args["expected_snapshot_hash"])
        post.assert_not_awaited()

    async def test_append_rejects_unreliable_inventory_in_preview_and_apply(self):
        for dry_run in (True, False):
            for inventory, code in (
                ({"items": [], "count": 1}, "ITEMS_INVENTORY_INCOMPLETE"),
                ({"items": [], "error_code": "ITEMS_UNAVAILABLE"}, "ITEMS_UNAVAILABLE"),
                ({"items": [{"layer": "0", "frame": 10, "length": 20}]}, "ITEMS_UNAVAILABLE"),
            ):
                with self.subTest(dry_run=dry_run, inventory=inventory):
                    responses = [inventory] if dry_run else [{"characters": [{"name": "A"}]}, inventory]
                    with patch.object(server, "ymm4_get", AsyncMock(side_effect=responses)), \
                         patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                        result = await server.dispatch({
                            "action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                            "append_after_existing": True, "dry_run": dry_run,
                        })
                    self.assertFalse(result["success"])
                    self.assertEqual(result["error_code"], code)
                    post.assert_not_awaited()

    async def test_guarded_append_uses_the_verified_inventory_for_placement(self):
        existing = {"item_id": "native:existing", "revision": "r1", "frame": 10, "layer": 0, "length": 20}
        snapshot = {"items": [existing], "count": 1}
        added = {"success": True, "item_id": "native:new", "revision": "r2",
                 "frame": 35, "layer": 0, "length": 15}
        responses = [snapshot, {"characters": [{"name": "A"}]}, snapshot,
                     {"items": [existing, added], "count": 2}]
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=responses)) as get, \
             patch.object(server, "ymm4_post", AsyncMock(return_value=added)) as post:
            result = await server.dispatch({
                "action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                "append_after_existing": True, "gap": 5,
                "expected_snapshot_hash": editplan.snapshot_hash([existing]),
            })
        self.assertTrue(result["success"])
        self.assertEqual(result["start_frame_used"], 35)
        self.assertEqual(post.await_args.args[1]["frame"], 35)
        self.assertEqual(get.await_args_list, [call("/items"), call("/characters"), call("/items"), call("/items")])

    async def test_guarded_append_preview_does_not_attest_to_a_changed_inventory(self):
        changed = {"items": [{"item_id": "native:other", "frame": 10, "layer": 0, "length": 20}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[{"items": []}, changed])), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({
                "action": "add_script", "lines": [{"text": "hello", "character": "A"}],
                "append_after_existing": True, "dry_run": True,
                "expected_snapshot_hash": editplan.snapshot_hash([]),
            })
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")
        self.assertNotIn("snapshot_hash", result)
        post.assert_not_awaited()
