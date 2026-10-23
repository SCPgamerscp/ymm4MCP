"""Missing assets are checked on the host, not on the MCP process machine."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class MissingAssetTests(unittest.IsolatedAsyncioTestCase):
    async def test_deduplicates_paths_and_reports_unknown_sources(self):
        items = {"success": True, "items": [
            {"item_id": "a", "type": "VideoItem", "source_path": "C:/Media/x.mp4"},
            {"item_id": "b", "type": "VideoItem", "source_path": "c:\\media\\X.mp4"},
            {"item_id": "c", "type": "AudioItem"},
            {"item_id": "d", "type": "TextItem"}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                items, {"success": True, "exists": False}])) as get:
            result = await server.dispatch({"action": "get_info", "sub_action": "missing_assets"})
        self.assertEqual(result["missing"], [{"path": "C:/Media/x.mp4", "item_ids": ["a", "b"]}])
        self.assertEqual(result["unknown_source_item_ids"], ["c"])
        self.assertFalse(result["complete"])
        self.assertEqual(get.await_count, 2)

    async def test_unavailable_host_does_not_claim_assets_present(self):
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                {"success": True, "items": [{"item_id": "a", "type": "ImageItem",
                                            "source_path": "C:/Media/a.png"}]},
                {"error": "unavailable"}])):
            result = await server.dispatch({"action": "get_info", "sub_action": "missing_assets"})
        self.assertEqual(result["checked_count"], 0)
        self.assertFalse(result["complete"])

    async def test_error_snapshot_with_items_is_rejected_before_probing_files(self):
        # A host error may still carry a stale items array and success=true.
        for error in ("timeline unavailable", None):
            with self.subTest(error=error):
                snapshot = {"success": True, "error": error, "items": [
                    {"item_id": "a", "type": "VideoItem",
                     "source_path": "C:/Media/a.mp4"}]}
                with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)) as get:
                    result = await server.dispatch({"action": "get_info",
                                                    "sub_action": "missing_assets"})
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                get.assert_awaited_once_with("/items")

    async def test_error_media_info_with_exists_true_is_not_trusted(self):
        snapshot = {"success": True, "items": [
            {"item_id": "a", "type": "ImageItem",
             "source_path": "C:/Media/a.png"}]}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                snapshot, {"success": True, "exists": True, "error": "access denied"}])) as get:
            result = await server.dispatch({"action": "get_info",
                                            "sub_action": "missing_assets"})
        self.assertFalse(result["complete"])
        self.assertEqual(result["checked_count"], 0)
        self.assertEqual(result["missing"], [])
        self.assertEqual(result["unchecked"], [
            {"path": "C:/Media/a.png", "item_ids": ["a"]}])
        self.assertEqual(get.await_count, 2)
