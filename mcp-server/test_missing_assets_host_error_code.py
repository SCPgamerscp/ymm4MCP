"""Asset QA must never turn contradictory host error responses into a pass."""
import unittest
from unittest.mock import AsyncMock, patch

import asset_qa
import server


class AssetHostErrorMarkerTests(unittest.IsolatedAsyncioTestCase):
    async def test_snapshot_error_code_or_missing_success_rejects_before_file_checks(self):
        snapshots = [
            {"success": True, "error_code": "ITEMS_UNAVAILABLE", "items": []},
            {"success": True, "error_code": None, "items": []},
            {"items": []},
        ]
        for snapshot in snapshots:
            with self.subTest(snapshot=snapshot):
                with patch.object(server, "ymm4_get", AsyncMock(return_value=snapshot)) as get:
                    result = await server.dispatch({
                        "action": "get_info", "sub_action": "missing_assets"})
                self.assertFalse(result["success"])
                self.assertEqual(result["error_code"], "ITEMS_UNAVAILABLE")
                get.assert_awaited_once_with("/items")

    async def test_media_info_error_code_never_counts_as_checked(self):
        snapshot = {"success": True, "items": [
            {"item_id": "img:1", "type": "ImageItem",
             "source_path": "C:/Media/picture.png"}]}
        for marker in ("READ_FAILED", None):
            with self.subTest(marker=marker):
                with patch.object(server, "ymm4_get", AsyncMock(side_effect=[
                        snapshot, {"success": True, "exists": True,
                                   "error_code": marker}])) as get:
                    result = await server.dispatch({
                        "action": "get_info", "sub_action": "missing_assets"})
                self.assertTrue(result["success"])
                self.assertFalse(result["complete"])
                self.assertEqual(result["checked_count"], 0)
                self.assertEqual(result["missing"], [])
                self.assertEqual(result["unchecked"], [
                    {"path": "C:/Media/picture.png", "item_ids": ["img:1"]}])
                self.assertEqual(get.await_count, 2)

    def test_asset_report_rejects_success_with_host_error(self):
        clean = {"success": True, "complete": True, "checked_count": 0,
                 "missing": [], "unknown_source_item_ids": [], "unchecked": []}
        self.assertTrue(asset_qa.report(clean)["passed"])
        for marker in ("error", "error_code"):
            for value in ("HOST_UNAVAILABLE", None):
                with self.subTest(marker=marker, value=value):
                    with self.assertRaisesRegex(ValueError, "asset availability check failed"):
                        asset_qa.report({**clean, marker: value})


if __name__ == "__main__":
    unittest.main()
