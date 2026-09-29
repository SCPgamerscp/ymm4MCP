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
