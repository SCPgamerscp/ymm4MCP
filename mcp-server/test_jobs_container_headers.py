"""Regression tests for bounded AVI/MOV/Matroska/WebM/WAV export header validation."""
import tempfile
import unittest
from pathlib import Path

import jobs


class ExportContainerHeaderTests(unittest.TestCase):
    def verify(self, name, data, fmt=None):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / name
            target.write_bytes(data.ljust(40, b"\0"))
            return jobs.verify_media_file(str(target), fmt)

    def assert_verified_header_only(self, result):
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["verification_scope"], "container_header_only")
        self.assertFalse(result["streams_verified"])

    def test_wav_riff_size_is_consistent_with_file(self):
        valid = b"RIFF" + (28).to_bytes(4, "little") + b"WAVE" + b"\0" * 24
        self.assert_verified_header_only(self.verify("audio.wav", valid))
        for declared in (0, 3, 0xffffffff):
            with self.subTest(declared=declared):
                corrupt = b"RIFF" + declared.to_bytes(4, "little") + b"WAVE" + b"\0" * 24
                self.assertEqual(self.verify("audio.wav", corrupt)["error_code"], "EXPORT_VERIFY_FAILED")

    def test_avi_requires_riff_avi_magic_and_bounded_size(self):
        valid = b"RIFF" + (28).to_bytes(4, "little") + b"AVI " + b"\0" * 24
        self.assert_verified_header_only(self.verify("clip.avi", valid))
        for corrupt in (b"RIFF" + (28).to_bytes(4, "little") + b"WAVE" + b"\0" * 24,
                        b"RIFF" + (9999).to_bytes(4, "little") + b"AVI " + b"\0" * 24,
                        b"garbage data" + b"\0" * 24):
            with self.subTest(corrupt=corrupt[:12]):
                self.assertFalse(self.verify("clip.avi", corrupt)["ok"])

    def test_mov_accepts_quicktime_ftyp_or_moov_only(self):
        ftyp = (16).to_bytes(4, "big") + b"ftypqt  " + b"\0" * 4
        moov = (8).to_bytes(4, "big") + b"moov"
        self.assert_verified_header_only(self.verify("clip.mov", ftyp))
        self.assert_verified_header_only(self.verify("clip.mov", moov))
        for invalid in ((4).to_bytes(4, "big") + b"moov",
                        (10000).to_bytes(4, "big") + b"moov",
                        (16).to_bytes(4, "big") + b"junkqt  " + b"\0" * 4):
            with self.subTest(invalid=invalid[:8]):
                self.assertFalse(self.verify("clip.mov", invalid)["ok"])

    def ebml(self, doctype):
        name = doctype.encode("ascii")
        element = b"\x42\x82" + bytes([0x80 + len(name)]) + name
        return b"\x1a\x45\xdf\xa3" + bytes([0x80 + len(element)]) + element

    def test_webm_and_matroska_must_match_ebml_doctype(self):
        self.assert_verified_header_only(self.verify("clip.webm", self.ebml("webm")))
        self.assert_verified_header_only(self.verify("clip.mkv", self.ebml("matroska")))
        self.assertFalse(self.verify("clip.webm", self.ebml("matroska"))["ok"])
        self.assertFalse(self.verify("clip.mkv", self.ebml("webm"))["ok"])
        self.assertFalse(self.verify("clip.webm", b"random" + b"\0" * 32)["ok"])

    def test_corrupt_ebml_header_is_rejected(self):
        for payload in (b"\x1a\x45\xdf\xa3\x00",
                        b"\x1a\x45\xdf\xa3\x40\xc8",
                        b"\x1a\x45\xdf\xa3\x84\x42\x82\x8awebm"):
            with self.subTest(payload=payload):
                self.assertFalse(self.verify("clip.webm", payload)["ok"])

    def test_ebml_void_payload_cannot_impersonate_doctype(self):
        fake = b"\x42\x82\x84junk"
        void = b"\xec" + bytes([0x80 + len(fake)]) + fake
        real = b"\x42\x82\x84webm"
        header = b"\x1a\x45\xdf\xa3"
        good = header + bytes([0x80 + len(void + real)]) + void + real
        bad = header + bytes([0x80 + len(void)]) + void
        self.assert_verified_header_only(self.verify("clip.webm", good))
        self.assertFalse(self.verify("clip.webm", bad)["ok"])

    def test_mov_finds_moov_after_free_skip_or_mdat(self):
        moov = (8).to_bytes(4, "big") + b"moov"
        for first_type in (b"free", b"skip", b"mdat"):
            with self.subTest(first_type=first_type):
                first = (16).to_bytes(4, "big") + first_type + b"payload!"
                self.assert_verified_header_only(self.verify("clip.mov", first + moov))
        # A marker inside mdat is not a top-level atom.
        payload = b"garbage!" + moov
        mdat = (8 + len(payload)).to_bytes(4, "big") + b"mdat" + payload
        self.assertFalse(self.verify("clip.mov", mdat)["ok"])

    def test_mov_rejects_atom_bounds_and_accepts_extended_size(self):
        moov = (8).to_bytes(4, "big") + b"moov"
        free = (16).to_bytes(4, "big") + b"free" + b"padding!"
        self.assert_verified_header_only(self.verify("clip.mov", free + moov))
        extended = (1).to_bytes(4, "big") + b"mdat" + (24).to_bytes(8, "big") + b"payload!"
        self.assert_verified_header_only(self.verify("clip.mov", extended + moov))
        corrupt = (10000).to_bytes(4, "big") + b"mdat" + b"payload!" + moov
        self.assertFalse(self.verify("clip.mov", corrupt)["ok"])

    def test_expected_format_must_match_extension_and_allowlist(self):
        valid = (16).to_bytes(4, "big") + b"ftypisom" + b"\0" * 4
        self.assertFalse(self.verify("clip.mp4", valid, "wav")["ok"])
        self.assertFalse(self.verify("clip.mp4", valid, "exe")["ok"])
        self.assertFalse(self.verify("clip.mp4", valid, True)["ok"])


if __name__ == "__main__":
    unittest.main()
