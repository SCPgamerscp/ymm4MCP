"""Asset discovery dispatch validation without YMM4."""
import unittest
from unittest.mock import AsyncMock, patch

import server


class AssetDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_encoded_query_and_options(self):
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "assets",
                                   "directory": "C:/素材/効果音", "query": "爆発&SE",
                                   "recursive": True, "hash": True, "max_results": 20})
            get.assert_awaited_once_with("/media/assets?directory=C%3A%2F%E7%B4%A0%E6%9D%90%2F%E5%8A%B9%E6%9E%9C%E9%9F%B3&query=%E7%88%86%E7%99%BA%26SE&recursive=true&hash=true&max_results=20")

    async def test_invalid_options_do_not_call_ymm4(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            for args in ({"directory": "relative"}, {"directory": "C:/media", "hash": "true"},
                         {"directory": "C:/media", "max_results": 0},
                         {"directory": "C:/media", "recursive": 1}):
                with self.subTest(args=args), self.assertRaises(ValueError):
                    await server.dispatch({"action": "get_info", "sub_action": "assets", **args})
            get.assert_not_awaited()
