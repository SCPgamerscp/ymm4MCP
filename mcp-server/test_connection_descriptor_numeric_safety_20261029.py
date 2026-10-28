"""Regression tests for bounded numeric JSON connection descriptors (Issue #22)."""
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import (
    ConnectionConfigurationError,
    MAX_DESCRIPTOR_FLOAT_CHARS,
    MAX_DESCRIPTOR_INTEGER_DIGITS,
    connection_settings,
)


class DescriptorNumericSafetyTests(unittest.TestCase):
    def settings_for(self, numeric_literal: bytes, token: str = ""):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "connection.json"
            path.write_bytes(
                b'{"token":"secret","ignored":{"numeric":' + numeric_literal + b"}}"
            )
            env = {"YMM4_CONNECTION_FILE": str(path)}
            if token:
                env["YMM4_API_TOKEN"] = token
            with patch.dict(os.environ, env, clear=True):
                return connection_settings()

    def test_positive_and_negative_exponent_overflow_are_rejected(self):
        for literal in (b"1e309", b"-1e309", b"9.99e9999"):
            with self.subTest(literal=literal):
                with self.assertRaises(ConnectionConfigurationError):
                    self.settings_for(literal)

    def test_float_overflow_inside_ignored_nested_metadata_is_rejected(self):
        with self.assertRaises(ConnectionConfigurationError):
            self.settings_for(b"[0,1e999]")

    def test_integer_over_digit_limit_is_rejected_before_conversion(self):
        for literal in (b"9" * (MAX_DESCRIPTOR_INTEGER_DIGITS + 1),
                        b"-" + b"9" * (MAX_DESCRIPTOR_INTEGER_DIGITS + 1)):
            with self.subTest(length=len(literal)):
                with self.assertRaises(ConnectionConfigurationError):
                    self.settings_for(literal)

    def test_extremely_large_integer_is_rejected_without_python_digit_exception(self):
        with self.assertRaises(ConnectionConfigurationError):
            self.settings_for(b"9" * 5000)

    def test_float_literal_over_length_limit_is_rejected(self):
        # 0. + many fractional digits is finite but still unnecessarily expensive.
        literal = b"0." + b"1" * MAX_DESCRIPTOR_FLOAT_CHARS
        with self.assertRaises(ConnectionConfigurationError):
            self.settings_for(literal)

    def test_valid_integer_at_digit_limit_is_accepted(self):
        literal = b"9" * MAX_DESCRIPTOR_INTEGER_DIGITS
        self.assertEqual(
            self.settings_for(literal)[1]["X-Ymm4-Token"], "secret"
        )

    def test_finite_float_and_exponent_are_accepted(self):
        for literal in (b"1e308", b"-1.25e-100", b"0.0", b"1.5"):
            with self.subTest(literal=literal):
                self.assertEqual(
                    self.settings_for(literal)[1]["X-Ymm4-Token"], "secret"
                )

    def test_invalid_numeric_descriptor_keeps_explicit_token_fallback(self):
        self.assertEqual(
            self.settings_for(b"1e999", token="explicit"),
            ("http://127.0.0.1:8765/api", {"X-Ymm4-Token": "explicit"}),
        )


if __name__ == "__main__":
    unittest.main()
