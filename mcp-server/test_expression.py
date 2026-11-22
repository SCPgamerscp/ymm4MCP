"""Explicit expression presets must target one current FaceItem without guessing."""
import unittest
from unittest.mock import AsyncMock, patch

import server


ARGS = {"action": "set_expression", "item_id": "native:face", "expression": "angry",
        "expression_map": {"angry": "C:/faces/angry.png"}}
ITEMS = {"success": True, "items": [{"item_id": "native:face", "type": "FaceItem",
                                     "revision": "r1"}]}
VERIFIED = {"success": True, "item_id": "native:face", "changed": True, "revision": "r2",
            "face_path": "C:/faces/angry.png", "results": ["FacePath=OK"],
            "error_code": None, "outcome_unknown": False}


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
             patch.object(server, "ymm4_post", AsyncMock(return_value=VERIFIED)) as post:
            result = await server.dispatch({**ARGS, "dry_run": False, "expected_revision": "r1"})
        self.assertTrue(result["success"])
        self.assertTrue(result["verified"])
        self.assertEqual(result["revision"], "r2")
        post.assert_awaited_once_with("/items/face/param", {
            "item_id": "native:face", "expected_revision": "r1", "FacePath": "C:/faces/angry.png",
            "require_face_path_readback": True})

    async def test_unverified_write_cannot_certify_expression_change(self):
        replies = [
            {"success": True, "revision": "r2"},
            {**VERIFIED, "item_id": "native:other"},
            {**VERIFIED, "revision": "r1"},
            {**VERIFIED, "revision": "  "},
            {**VERIFIED, "face_path": "C:/faces/happy.png"},
            {**VERIFIED, "results": ["FacePath=NG"]},
            {**VERIFIED, "error_code": "FACE_PARAM_UNVERIFIED"},
            {**VERIFIED, "changed": False},
            {"success": False, "outcome_unknown": True, "error_code": "FACE_PARAM_UNVERIFIED"},
        ]
        for reply in replies:
            with self.subTest(reply=reply), \
                 patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)), \
                 patch.object(server, "ymm4_post", AsyncMock(return_value=reply)) as post:
                result = await server.dispatch({**ARGS, "dry_run": False})
            self.assertFalse(result["success"])
            self.assertEqual(result["error_code"], "EXPRESSION_OUTCOME_UNKNOWN")
            self.assertTrue(result["outcome_unknown"])
            self.assertNotIn("verified", result)
            post.assert_awaited_once()

    async def test_preflight_rejection_on_older_plugin_is_reported_as_failed(self):
        reply = {"success": False, "changed": False, "error_code": "FACE_PARAM_INVALID",
                 "results": ["require_face_path_readback=NOTFOUND"]}
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)), \
             patch.object(server, "ymm4_post", AsyncMock(return_value=reply)):
            result = await server.dispatch({**ARGS, "dry_run": False})
        self.assertEqual(result["error_code"], "EXPRESSION_UPDATE_FAILED")
        self.assertNotIn("outcome_unknown", result)

    async def test_invalid_inventory_and_foreign_type_never_mutate(self):
        other = {"items": [{**ITEMS["items"][0], "type": "Unrelated.NotFaceItem"}]}
        for snapshot, code in ((other, "FACE_ITEM_NOT_FOUND"),
                               ({**ITEMS, "count": 2}, "ITEMS_INVENTORY_INCOMPLETE"),
                               ({**ITEMS, "error_code": "UNAVAILABLE"}, "ITEMS_UNAVAILABLE")):
            with self.subTest(snapshot=snapshot), \
                 patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)), \
                 patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
                result = await server.dispatch({**ARGS, "dry_run": False})
            self.assertEqual(result["error_code"], code)
            post.assert_not_awaited()

    async def test_stale_revision_and_invalid_mapping_never_mutate(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value=ITEMS)), \
             patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            result = await server.dispatch({**ARGS, "dry_run": False, "expected_revision": "old"})
            self.assertEqual(result["error_code"], "REVISION_CONFLICT")
            with self.assertRaises(ValueError):
                await server.dispatch({**ARGS, "expression_map": {"happy": "unused"}})
            post.assert_not_awaited()
