"""Snapshot verification must reject untrustworthy host/reference evidence."""
import unittest
from unittest.mock import AsyncMock

from quality_guard import MAX_SNAPSHOT_ITEMS, verify_snapshot

GOOD = [{"item_id": "voice:1", "frame": 0, "length": 30, "endFrame": 30,
         "layer": 1, "revision": "r1", "type": "VoiceItem",
         "text": "こんにちは\n世界", "source_path": None}]


class SnapshotEvidenceTests(unittest.IsolatedAsyncioTestCase):
    async def test_valid_endframe_and_legacy_omission(self):
        self.assertIsNone(await verify_snapshot(
            AsyncMock(return_value={"items": GOOD}), GOOD))
        legacy = [{"item_id": "voice:1", "frame": 0, "length": 30}]
        self.assertIsNone(await verify_snapshot(
            AsyncMock(return_value={"items": legacy}), legacy))

    async def test_reference_contradictory_endframe_stops_before_host(self):
        for value in (0, 29, 31, True, "30", -1):
            with self.subTest(value=value):
                invalid = [{**GOOD[0], "endFrame": value}]
                get = AsyncMock()
                result = await verify_snapshot(get, invalid)
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                get.assert_not_awaited()

    async def test_host_contradictory_endframe_cannot_pass(self):
        invalid = [{**GOOD[0], "endFrame": 29}]
        result = await verify_snapshot(AsyncMock(return_value={"items": invalid}), GOOD)
        self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")

    async def test_text_and_source_path_shapes_rejected(self):
        for field, invalids in (("text", ({"bad": "object"}, True, "a" * 65_537,
                                         "a\ud800", "a\u0085")),
                                ("source_path", ({"bad": "object"}, ["C:/file.mp4"],
                                                 True, "a" * 4097, "a\u202e", "a\ud800"))):
            for value in invalids:
                with self.subTest(field=field, type=type(value).__name__):
                    get = AsyncMock()
                    result = await verify_snapshot(get, [{**GOOD[0], field: value}])
                    self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                    get.assert_not_awaited()

    async def test_bad_host_text_and_path_cannot_match_reference(self):
        for field, value in (("text", {"bad": "object"}),
                             ("source_path", "C:/file\u202e.mp4")):
            invalid = [{**GOOD[0], field: value}]
            result = await verify_snapshot(AsyncMock(return_value={"items": invalid}), GOOD)
            self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")

    async def test_revision_and_type_unicode_spoofing_rejected(self):
        for field in ("revision", "type"):
            for value in ("valid\u202e", "valid\u0085", "valid\ud800", "valid\ufdd0"):
                with self.subTest(field=field, value=repr(value)):
                    get = AsyncMock()
                    result = await verify_snapshot(get, [{**GOOD[0], field: value}])
                    self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                    get.assert_not_awaited()

    async def test_oversized_snapshot_stops_before_hash_or_host(self):
        get = AsyncMock()
        rows = [{"item_id": "same"}] * (MAX_SNAPSHOT_ITEMS + 1)
        result = await verify_snapshot(get, rows)
        self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
        get.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
