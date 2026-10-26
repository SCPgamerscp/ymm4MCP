"""Snapshot QA must not pass equally malformed host/reference metadata."""
import unittest
from unittest.mock import AsyncMock

from quality_guard import verify_snapshot


GOOD = [{"item_id": "voice:1", "frame": 0, "length": 30, "layer": 1,
         "revision": "r1", "type": "VoiceItem"}]


class ObservableFieldTests(unittest.IsolatedAsyncioTestCase):
    async def test_valid_snapshots_and_legacy_optional_fields(self):
        self.assertIsNone(await verify_snapshot(
            AsyncMock(return_value={"success": True, "items": GOOD}), GOOD))
        legacy = [{"item_id": "voice:1"}]
        self.assertIsNone(await verify_snapshot(
            AsyncMock(return_value={"items": legacy}), legacy))

    async def test_invalid_reference_geometry_never_reads_host(self):
        for field, values in (
            ("frame", (True, "0", -1, 2_147_483_648, 1.5)),
            ("layer", (False, "1", -1, 2_147_483_648)),
            ("length", (True, 0, -1, "30", 2_147_483_648)),
        ):
            for value in values:
                with self.subTest(field=field, value=value):
                    invalid = [{**GOOD[0], field: value}]
                    get = AsyncMock()
                    result = await verify_snapshot(get, invalid)
                    self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                    get.assert_not_awaited()

    async def test_invalid_host_geometry_cannot_match_reference(self):
        for field, value in (("frame", True), ("layer", "1"), ("length", 0)):
            with self.subTest(field=field):
                invalid = [{**GOOD[0], field: value}]
                result = await verify_snapshot(
                    AsyncMock(return_value={"items": invalid}), invalid)
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")

    async def test_invalid_reference_revision_or_type_never_reads_host(self):
        for field in ("revision", "type"):
            for value in ("", " ", None, True, 1, "a\n", "x" * 257):
                with self.subTest(field=field, value=repr(value)):
                    invalid = [{**GOOD[0], field: value}]
                    get = AsyncMock()
                    result = await verify_snapshot(get, invalid)
                    self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")
                    get.assert_not_awaited()

    async def test_invalid_host_revision_or_type_cannot_pass(self):
        for field, value in (("revision", None), ("revision", ""), ("type", False)):
            with self.subTest(field=field):
                invalid = [{**GOOD[0], field: value}]
                result = await verify_snapshot(
                    AsyncMock(return_value={"success": True, "items": invalid}), GOOD)
                self.assertEqual(result["error_code"], "SNAPSHOT_VERIFY_UNKNOWN")

    async def test_well_formed_conflict_still_detected(self):
        changed = [{**GOOD[0], "frame": 5}]
        result = await verify_snapshot(
            AsyncMock(return_value={"success": True, "items": changed}), GOOD)
        self.assertEqual(result["error_code"], "SNAPSHOT_CONFLICT")


if __name__ == "__main__":
    unittest.main()
