"""Real /api/items endFrame geometry regressions for ripple verification."""
import unittest

from ripple_verification import verify_ripple_result


def snapshot(*items):
    return {"success": True, "count": len(items), "items": list(items)}


BEFORE = snapshot(
    {"item_id": "cut", "frame": 10, "length": 10, "endFrame": 20,
     "layer": 1, "revision": "r1", "text": "remove"},
    {"item_id": "later", "frame": 20, "length": 10, "endFrame": 30,
     "layer": 1, "revision": "r2", "text": "preserve"})
AFTER = snapshot(
    {"item_id": "later", "frame": 10, "length": 10, "endFrame": 20,
     "layer": 1, "revision": "r3", "text": "preserve"})
REVISIONS = {"cut": "r1", "later": "r2"}


def verify(after, before=BEFORE):
    return verify_ripple_result(before, after, start_frame=10, end_frame=20,
                                layers=[1], expected_revisions=REVISIONS)


class HostEndFrameTests(unittest.TestCase):
    def test_real_host_camelcase_endframe_moves_with_item(self):
        result = verify(AFTER)
        self.assertTrue(result["passed"], result["issues"])

    def test_stale_endframe_rejected(self):
        altered = snapshot(dict(AFTER["items"][0], endFrame=30))
        result = verify(altered)
        self.assertFalse(result["passed"])
        self.assertEqual(result["error_code"], "RIPPLE_AFTER_SNAPSHOT_INVALID")

    def test_missing_endframe_not_accepted_as_same_content(self):
        altered = snapshot({k: v for k, v in AFTER["items"][0].items()
                            if k != "endFrame"})
        result = verify(altered)
        self.assertFalse(result["passed"])
        self.assertTrue(any(x["field"] == "endFrame" for x in result["issues"]))

    def test_unrelated_text_change_still_rejected(self):
        altered = snapshot(dict(AFTER["items"][0], text="corrupted"))
        result = verify(altered)
        self.assertFalse(result["passed"])
        self.assertTrue(any(x["code"] == "RIPPLE_CONTENT_CHANGED"
                            and x["field"] == "text" for x in result["issues"]))

    def test_snake_case_end_frame_remains_supported(self):
        before = snapshot(*(dict({k: v for k, v in item.items()
                                  if k != "endFrame"}, end_frame=item["endFrame"])
                            for item in BEFORE["items"]))
        after = snapshot(dict({k: v for k, v in AFTER["items"][0].items()
                               if k != "endFrame"}, end_frame=20))
        self.assertTrue(verify(after, before)["passed"])


if __name__ == "__main__":
    unittest.main()
