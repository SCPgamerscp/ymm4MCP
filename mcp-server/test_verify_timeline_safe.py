"""Regression tests for read-only authenticated timeline QA."""
import importlib.util
from pathlib import Path
from unittest import TestCase, mock
import unittest

SOURCE = Path(__file__).resolve().parents[1] / "verify_timeline.py"
spec = importlib.util.spec_from_file_location("safe_timeline_verifier", SOURCE)
verifier = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verifier)


def item(item_id, frame, length, layer=0, text=None, kind="VoiceItem"):
    return {"item_id": item_id, "frame": frame, "length": length,
            "layer": layer, "type": kind, "text": text}


class TimelineVerifierTests(TestCase):
    def test_repeated_dialogue_is_informational(self):
        report = verifier.inspect([item("a", 0, 10, text="hello"),
                                   item("b", 10, 10, text="hello")])
        self.assertTrue(report["passed"])
        self.assertEqual(1, len(report["repeated_dialogue"]))
        self.assertEqual(0, report["overlap_count"])

    def test_nested_overlaps_are_not_hidden(self):
        report = verifier.inspect([item("a", 0, 100), item("b", 10, 2),
                                   item("c", 20, 2)])
        self.assertFalse(report["passed"])
        self.assertEqual(2, report["overlap_count"])
        self.assertEqual([["a", "b"], ["a", "c"]],
                         [issue["item_ids"] for issue in report["issues"]])

    def test_touching_edges_are_not_overlaps(self):
        self.assertTrue(verifier.inspect([item("a", 0, 10),
                                          item("b", 10, 5)])["passed"])

    def test_different_layers_are_independent(self):
        self.assertTrue(verifier.inspect([item("a", 0, 20),
                                          item("b", 5, 20, layer=1)])["passed"])

    def test_missing_text_is_valid_for_media(self):
        self.assertTrue(verifier.inspect([item("a", 0, 20, kind="VideoItem")])["passed"])

    def test_reject_duplicate_ids(self):
        with self.assertRaises(ValueError):
            verifier.inspect([item("a", 0, 10), item("a", 10, 10)])

    def test_reject_invalid_numeric_fields(self):
        with self.assertRaises(ValueError):
            verifier.inspect([item("a", False, 10)])
        with self.assertRaises(ValueError):
            verifier.inspect([item("a", 0, 0)])

    def test_reject_frame_overflow(self):
        with self.assertRaises(ValueError):
            verifier.inspect([item("a", 2147483640, 10)])

    def test_nonzero_exit_for_qa_failure(self):
        # Mock the async operation, not asyncio.run: mocking run discards the
        # verify() coroutine without awaiting it and hides async integration bugs.
        with mock.patch.object(verifier, "verify", new_callable=mock.AsyncMock,
                               return_value={"passed": False}) as verify:
            with mock.patch("builtins.print"):
                self.assertEqual(2, verifier.main())
            verify.assert_awaited_once_with()

    def test_nonzero_exit_for_connection_failure(self):
        with mock.patch.object(verifier, "verify", new_callable=mock.AsyncMock,
                               side_effect=ValueError("bad response")) as verify:
            with mock.patch("builtins.print"):
                self.assertEqual(1, verifier.main())
            verify.assert_awaited_once_with()


if __name__ == "__main__":
    unittest.main()
