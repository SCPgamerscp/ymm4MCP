"""Explicit expression presets must target one current FaceItem without guessing."""
import unittest
from unittest.mock import AsyncMock, patch

import server


ARGS = {"action": "set_expression", "item_id": "native:face", "expression": "angry",
        "expression_map": {"angry": "C:/faces/angry.png"}}
ITEMS = {"success": True, "items": [{"item_id": "native:face", "type": "FaceItem",
                                     "revision": "r1"}]}


class ExpressionTests(unittest.IsolatedAsyncioTestCase):
    async def test_dry_run_resolves_preset_without_mutation(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)) as get, \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch(ARGS)
        self.assertTrue(result["dry_run"])
        self.assertEqual(result["face_path"], "C:/faces/angry.png")
        get.assert_awaited_once_with("/items")
        post.assert_not_awaited()

    async def test_apply_uses_current_revision(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)), \
             patch.object(server, "ymm4_post", AsyncMock(return_value={
                 "success": True, "revision": "r2"})) as post:
            result = await server.dispatch({**ARGS, "dry_run": False, "expected_revision": "r1"})
        self.assertTrue(result["success"])
        self.assertEqual(result["revision"], "r2")
        post.assert_awaited_once_with("/items/face/param", {
            "item_id": "native:face", "expected_revision": "r1", "FacePath": "C:/faces/angry.png"})

    async def test_stale_revision_and_invalid_mapping_never_mutate(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({**ARGS, "dry_run": False, "expected_revision": "old"})
            self.assertEqual(result["error_code"], "REVISION_CONFLICT")
            with self.assertRaises(ValueError):
                await server.dispatch({**ARGS, "expression_map": {"happy": "unused"}})
            post.assert_not_awaited()
