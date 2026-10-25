"""Snapshot hashes must not certify ambiguous or malformed host timelines."""

import unittest

import editplan


class SnapshotHashIntegrityTests(unittest.TestCase):
    def test_valid_timeline_is_order_independent(self):
        items = [
            {"item_id": "voice:霊夢", "revision": "r1", "frame": 0, "length": 20},
            {"item_id": "bgm:1", "revision": "r2", "frame": 10, "length": 100},
        ]
        self.assertEqual(editplan.snapshot_hash(items),
                         editplan.snapshot_hash(list(reversed(items))))
        self.assertNotEqual(editplan.snapshot_hash(items),
                            editplan.snapshot_hash([{**items[0], "revision": "r3"}, items[1]]))

    def test_duplicate_stable_id_is_rejected_even_when_other_fields_differ(self):
        items = [
            {"item_id": "same", "revision": "r1", "frame": 0},
            {"item_id": "same", "revision": "r2", "frame": 20},
        ]
        with self.assertRaisesRegex(ValueError, "duplicate item_id"):
            editplan.snapshot_hash(items)

    def test_missing_or_invalid_stable_id_is_rejected(self):
        for item in (
            {}, {"item_id": None}, {"item_id": 1},
            {"item_id": ""}, {"item_id": "  "},
            {"item_id": "line\\n2".replace("\\n", "\n")},
            {"item_id": "a\\x00b".replace("\\x00", "\x00")},
            {"item_id": "a\\x7fb".replace("\\x7f", "\x7f")},
        ):
            with self.subTest(item=item), self.assertRaises(ValueError):
                editplan.snapshot_hash([item])

    def test_nonfinite_host_numbers_are_rejected(self):
        for value in (float("nan"), float("inf"), float("-inf")):
            for field in ("frame", "layer", "length"):
                with self.subTest(value=value, field=field), self.assertRaisesRegex(
                        ValueError, "invalid or non-finite"):
                    editplan.snapshot_hash([{"item_id": "a", field: value}])

    def test_non_json_host_values_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "invalid or non-finite"):
            editplan.snapshot_hash([{"item_id": "a", "source_path": object()}])


if __name__ == "__main__":
    unittest.main()
