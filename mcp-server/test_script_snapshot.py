"""Guard script placement against edits between preview and apply."""
import unittest
from unittest.mock import AsyncMock, patch

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
