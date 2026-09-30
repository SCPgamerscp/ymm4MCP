"""A final QA gate must not pass with missing or unverified media."""
import unittest
from unittest.mock import AsyncMock, patch

import asset_qa
import server


class AssetQaTests(unittest.IsolatedAsyncioTestCase):
    def test_missing_and_unknown_sources_are_errors(self):
        result = asset_qa.report({"success": True, "complete": False, "checked_count": 1,
                                  "missing": [{"path": "C:/lost.mp4", "item_ids": ["v"]}],
                                  "unknown_source_item_ids": ["a"], "unchecked": []})
        self.assertFalse(result["passed"])
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["ASSET_MISSING", "ASSET_SOURCE_UNKNOWN"])

    async def test_gate_checks_host_paths_and_blocks_missing_file(self):
        items = {"success": True, "items": [{"item_id": "v", "type": "VideoItem",
                                            "source_path": "C:/lost.mp4", "frame": 0,
                                            "length": 10, "layer": 1}]}
        async def get(path):
            if path == "/items":
                return items
            self.assertEqual(path, "/media/info?path=C%3A%2Flost.mp4")
            return {"success": True, "exists": False}
        with patch.object(server, "ymm4_get", AsyncMock(side_effect=get)):
            result = await server.dispatch({"action": "qa_gate", "assets_check": {},
                                            "include_gaps": False})
        self.assertEqual(result["decision"], "repair")
        self.assertEqual(result["qa"]["issues"][0]["code"], "ASSET_MISSING")
