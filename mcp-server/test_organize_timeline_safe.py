"""Regression tests for the safe standalone timeline organizer."""
import asyncio
import copy
import importlib.util
from pathlib import Path
import unittest

SOURCE = Path(__file__).resolve().parents[1] / "organize_timeline.py"
spec = importlib.util.spec_from_file_location("safe_timeline_organizer", SOURCE)
organizer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(organizer)


def item(item_id, layer, frame, length, text="same"):
    return {"item_id": item_id, "revision": "r1", "layer": layer,
            "frame": frame, "length": length, "type": "VoiceItem", "text": text}


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return copy.deepcopy(self.payload)


class FakeClient:
    def __init__(self, items, stale=False, unverified=False):
        self.items = copy.deepcopy(items)
        self.stale = stale
        self.unverified = unverified
        self.get_calls = 0
        self.post_calls = 0

    async def get(self, url, headers):
        self.get_calls += 1
        if self.stale and self.get_calls == 1:
            self.items[0]["frame"] += 1
        return FakeResponse({"items": self.items})

    async def post(self, url, headers, json):
        self.post_calls += 1
        if self.unverified:
            return FakeResponse({"success": True, "verified": False})
        target = next(x for x in self.items if x["item_id"] == json["item_id"])
        assert target["revision"] == json["expected_revision"]
        target["frame"] = int(json["value"])
        target["revision"] = "r2"
        return FakeResponse({"success": True, "verified": True,
                             "item_id": target["item_id"], "revision": "r2"})


class OrganizerPlanTests(unittest.TestCase):
    def test_repeated_dialogue_is_never_deleted(self):
        items = [item("a", 0, 0, 10), item("b", 0, 10, 10)]
        self.assertEqual([], organizer.plan_moves(items))
        self.assertEqual(2, len(items))

    def test_cascading_overlaps(self):
        moves = organizer.plan_moves([item("a", 0, 0, 10), item("b", 0, 5, 8),
                                      item("c", 0, 10, 4)])
        self.assertEqual([("b", 10), ("c", 18)],
                         [(m["item_id"], m["new_frame"]) for m in moves])

    def test_layers_are_independent(self):
        moves = organizer.plan_moves([item("a", 0, 0, 10), item("b", 1, 5, 10),
                                      item("c", 0, 3, 2)])
        self.assertEqual([("c", 10)], [(m["item_id"], m["new_frame"]) for m in moves])

    def test_reject_duplicate_and_missing_revision(self):
        with self.assertRaises(ValueError):
            organizer.plan_moves([item("a", 0, 0, 10), item("a", 0, 5, 10)])
        bad = item("a", 0, 0, 10)
        del bad["revision"]
        with self.assertRaises(ValueError):
            organizer.plan_moves([bad])

    def test_reject_boolean_and_overflow(self):
        with self.assertRaises(ValueError):
            organizer.plan_moves([item("a", 0, False, 10)])
        with self.assertRaises(ValueError):
            organizer.plan_moves([item("a", 0, 2147483640, 10)])
        with self.assertRaises(ValueError):
            organizer.plan_moves([item("a", 0, 2147483630, 10),
                                  item("b", 0, 2147483631, 10)])


class OrganizerApplyTests(unittest.IsolatedAsyncioTestCase):
    async def test_verified_move(self):
        items = [item("a", 0, 0, 10), item("b", 0, 5, 10)]
        client = FakeClient(items)
        done = await organizer.apply_moves(client, "http://127.0.0.1:8765/api",
                                           {"X-Ymm4-Token": "secret"},
                                           items, organizer.plan_moves(items))
        self.assertEqual(1, done)
        self.assertEqual(10, client.items[1]["frame"])
        self.assertEqual(1, client.post_calls)

    async def test_stale_snapshot_prevents_all_writes(self):
        items = [item("a", 0, 0, 10), item("b", 0, 5, 10)]
        client = FakeClient(items, stale=True)
        with self.assertRaises(RuntimeError):
            await organizer.apply_moves(client, "http://127.0.0.1:8765/api", {},
                                        items, organizer.plan_moves(items))
        self.assertEqual(0, client.post_calls)

    async def test_unverified_write_stops(self):
        items = [item("a", 0, 0, 10), item("b", 0, 5, 10)]
        client = FakeClient(items, unverified=True)
        with self.assertRaises(RuntimeError):
            await organizer.apply_moves(client, "http://127.0.0.1:8765/api", {},
                                        items, organizer.plan_moves(items))
        self.assertEqual(1, client.post_calls)


if __name__ == "__main__":
    unittest.main()
