"""Regressions for stable-ID post-edit verification (Issue #61)."""
import unittest
from unittest.mock import AsyncMock, patch

from item_verification import validate_expected_items, verify_items


def snapshot(*items):
    return {"items": list(items), "count": len(items)}


def item(item_id="native:a", **overrides):
    result = {"item_id": item_id, "revision": "r1", "frame": 10, "layer": 2,
              "length": 50, "type": "ImageItem", "text": ""}
    result.update(overrides)
    return result


class ItemVerificationTests(unittest.TestCase):
    def test_matching_fields_pass(self):
        result = verify_items(snapshot(item()), [{"item_id": "native:a", "revision": "r1",
                                                    "frame": 10, "layer": 2, "length": 50}])
        self.assertTrue(result["passed"])
        self.assertEqual(result["matched_count"], 1)

    def test_field_mismatch_identifies_field(self):
        result = verify_items(snapshot(item(frame=11, revision="r2")),
                              [{"item_id": "native:a", "frame": 10, "revision": "r1"}])
        self.assertFalse(result["passed"])
        self.assertEqual({issue["field"] for issue in result["issues"]}, {"frame", "revision"})

    def test_missing_item_fails(self):
        result = verify_items(snapshot(item()), [{"item_id": "native:missing", "frame": 10}])
        self.assertEqual(result["issues"][0]["code"], "ITEM_MISSING")

    def test_duplicate_item_id_fails_even_when_one_matches(self):
        result = verify_items(snapshot(item(), item(frame=99)),
                              [{"item_id": "native:a", "frame": 10}])
        self.assertEqual(result["error_code"], "ITEM_ID_AMBIGUOUS")

    def test_unrelated_duplicate_fails_closed(self):
        result = verify_items(snapshot(item(), item("native:b"), item("native:b")),
                              [{"item_id": "native:a", "frame": 10}])
        self.assertFalse(result["success"])

    def test_missing_field_is_not_equal_to_null(self):
        actual = item()
        del actual["revision"]
        result = verify_items(snapshot(actual), [{"item_id": "native:a", "revision": "r1"}])
        self.assertEqual(result["issues"][0]["code"], "ITEM_FIELD_UNAVAILABLE")

    def test_snapshot_count_mismatch_rejected(self):
        bad = snapshot(item())
        bad["count"] = 2
        self.assertEqual(verify_items(bad, [{"item_id": "native:a", "frame": 10}])["error_code"],
                         "ITEMS_SNAPSHOT_INVALID")

    def test_snapshot_failure_rejected(self):
        self.assertEqual(verify_items({"error": "host unavailable", "items": [], "count": 0},
                                      [{"item_id": "native:a", "frame": 10}])["error_code"],
                         "ITEMS_UNAVAILABLE")

    def test_invalid_item_in_snapshot_rejected(self):
        self.assertEqual(verify_items(snapshot(item(), {"frame": 2}),
                                      [{"item_id": "native:a", "frame": 10}])["error_code"],
                         "ITEMS_SNAPSHOT_INVALID")

    def test_bool_not_accepted_as_frame(self):
        with self.assertRaises(ValueError):
            validate_expected_items([{"item_id": "native:a", "frame": True}])

    def test_repeated_expected_id_rejected(self):
        with self.assertRaises(ValueError):
            validate_expected_items([{"item_id": "native:a", "frame": 1},
                                     {"item_id": "native:a", "frame": 2}])

    def test_empty_and_id_only_expectations_rejected(self):
        for expected in ([], [{"item_id": "native:a"}]):
            with self.subTest(expected=expected), self.assertRaises(ValueError):
                validate_expected_items(expected)

    def test_unexpected_field_rejected(self):
        with self.assertRaises(ValueError):
            validate_expected_items([{"item_id": "native:a", "unsupported": 1}])

    def test_field_type_mismatch_does_not_pass(self):
        result = verify_items(snapshot(item(frame=True)),
                              [{"item_id": "native:a", "frame": 1}])
        self.assertFalse(result["passed"])

    def test_multiple_expected_items_partial_match(self):
        result = verify_items(snapshot(item(), item("native:b", frame=20)),
                              [{"item_id": "native:a", "frame": 10},
                               {"item_id": "native:b", "frame": 30}])
        self.assertEqual(result["matched_count"], 1)
        self.assertFalse(result["passed"])


class ItemVerificationDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_dispatch_uses_one_read_only_snapshot(self):
        import server
        with patch.object(server, "ymm4_get", new_callable=AsyncMock,
                          return_value=snapshot(item())) as get:
            result = await server.dispatch({"action": "verify_items",
                                            "expected": [{"item_id": "native:a", "frame": 10}]})
        self.assertTrue(result["passed"])
        get.assert_awaited_once_with("/items")

    async def test_dispatch_rejects_invalid_expected_before_http(self):
        import server
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "verify_items", "expected": []})
        get.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
