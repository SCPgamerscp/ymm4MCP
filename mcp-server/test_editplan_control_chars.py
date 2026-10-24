"""Reject invisible control characters in EditPlan identity fields."""

import unittest

import editplan


def minimal_plan():
    return {"plan": {"scenes": [{"id": "scene-1", "items": [
        {"id": "caption_1", "type": "text", "text": "hello", "length": 10}
    ]}]}, "idempotency_key": "episode-1"}


class EditPlanControlCharacterTests(unittest.TestCase):
    def test_rejects_nonprinting_scene_ids(self):
        for char in ("\x00", "\x01", "\x1b", "\x7f"):
            with self.subTest(char=repr(char)):
                args = minimal_plan()
                args["plan"]["scenes"][0]["id"] = "scene" + char + "1"
                with self.assertRaises(ValueError):
                    editplan.parse_plan(args)

    def test_rejects_nonprinting_item_ids(self):
        for char in ("\x00", "\x1f", "\x7f"):
            with self.subTest(char=repr(char)):
                args = minimal_plan()
                args["plan"]["scenes"][0]["items"][0]["id"] = "caption" + char + "1"
                with self.assertRaises(ValueError):
                    editplan.parse_plan(args)

    def test_rejects_control_characters_inside_idempotency_keys(self):
        for char in ("\x00", "\x09", "\x1b", "\x7f"):
            with self.subTest(char=repr(char)):
                args = minimal_plan()
                args["idempotency_key"] = "episode" + char + "1"
                with self.assertRaises(ValueError):
                    editplan.parse_plan(args)

    def test_accepts_printable_ids_and_unicode_idempotency_keys(self):
        args = minimal_plan()
        args["plan"]["scenes"][0]["id"] = "scene:1"
        args["plan"]["scenes"][0]["items"][0]["id"] = "caption@1"
        args["idempotency_key"] = "動画-episode 1"
        parsed = editplan.parse_plan(args)
        self.assertEqual(parsed["scenes"][0]["id"], "scene:1")
        self.assertEqual(parsed["scenes"][0]["items"][0]["id"], "caption@1")
        self.assertEqual(parsed["idempotency_key"], "動画-episode 1")


if __name__ == "__main__":
    unittest.main()
