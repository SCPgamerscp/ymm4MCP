"""Reject ambiguous or malformed native additions during EditPlan verification."""
import unittest

import editplan


def detail(item_id="native:1", **overrides):
    result = {"revision": "r1", "frame": 0, "layer": 1, "length": 20}
    result.update(overrides)
    return {"id": "caption", "op": "add", "item_id": item_id, "result": result}


def current(item_id="native:1", **overrides):
    item = {"item_id": item_id, "revision": "r1", "frame": 0,
            "layer": 1, "length": 20}
    item.update(overrides)
    return item


class AppliedItemIntegrityTests(unittest.TestCase):
    def test_valid_addition_still_verifies(self):
        self.assertEqual(editplan.verify_applied_items([detail()], [current()]), [])

    def test_rejects_unhashable_or_blank_claimed_item_ids(self):
        for bad in ([], {}, None, ""):
            with self.subTest(bad=repr(bad)):
                failures = editplan.verify_applied_items([detail(bad)], [current()])
                self.assertEqual(failures[0]["reason"], "ITEM_NOT_UNIQUE")

    def test_two_additions_cannot_claim_the_same_native_item(self):
        claims = [detail(), {**detail(), "id": "caption-2"}]
        failures = editplan.verify_applied_items(claims, [current()])
        self.assertEqual(failures, [{"id": "caption-2", "item_id": "native:1",
                                     "reason": "ITEM_REUSED"}])

    def test_bool_and_float_numeric_readbacks_never_match_integers(self):
        for field in ("frame", "layer", "length"):
            expected_value = detail()["result"][field]
            for invalid in (True, float(expected_value)):
                with self.subTest(field=field, invalid=invalid):
                    failures = editplan.verify_applied_items([detail()], [
                        current(**{field: invalid})])
                    self.assertEqual(failures[0]["reason"], "STATE_MISMATCH")
                    self.assertIn(field, failures[0]["fields"])

    def test_bool_or_float_expected_values_do_not_match_host_integers(self):
        for field in ("frame", "layer", "length"):
            expected_value = detail()["result"][field]
            for invalid in (True, float(expected_value)):
                with self.subTest(field=field, invalid=invalid):
                    failures = editplan.verify_applied_items([
                        detail(**{field: invalid})], [current()])
                    self.assertIn(field, failures[0]["fields"])

    def test_revision_must_be_nonempty_text_on_both_sides(self):
        for bad in (1, True, "", None):
            with self.subTest(bad=bad):
                self.assertIn("revision", editplan.verify_applied_items([
                    detail()], [current(revision=bad)])[0]["fields"])
                self.assertIn("revision", editplan.verify_applied_items([
                    detail(revision=bad)], [current()])[0]["fields"])

    def test_non_dict_add_result_fails_closed(self):
        malformed = {**detail(), "result": None}
        failures = editplan.verify_applied_items([malformed], [current()])
        self.assertEqual(failures[0]["reason"], "STATE_MISMATCH")


if __name__ == "__main__":
    unittest.main()
