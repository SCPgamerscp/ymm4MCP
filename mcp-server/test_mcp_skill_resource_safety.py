"""Bundled MCP skill resources must stay within their trusted size and root."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import mcp_skills


class SkillResourceSafetyTests(unittest.TestCase):
    def test_rejects_skill_file_symlink_outside_bundle(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            root = base / "skills"
            target = base / "outside.md"
            target.write_text("untrusted content", encoding="utf-8")
            (root / "ymm4-jikkyou").mkdir(parents=True)
            (root / "ymm4-jikkyou" / "SKILL.md").symlink_to(target)
            with patch.object(mcp_skills, "ROOT", root):
                with self.assertRaisesRegex(ValueError, "escapes"):
                    mcp_skills.skill_text("jikkyou")

    def test_rejects_skill_directory_symlink_outside_bundle(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            root = base / "skills"
            root.mkdir()
            external = base / "external"
            external.mkdir()
            (external / "SKILL.md").write_text("untrusted", encoding="utf-8")
            (root / "ymm4-jikkyou").symlink_to(external, target_is_directory=True)
            with patch.object(mcp_skills, "ROOT", root):
                with self.assertRaisesRegex(ValueError, "escapes"):
                    mcp_skills.skill_text("jikkyou")

    def test_rejects_oversized_skill_without_reading_entire_file(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            skill = root / "ymm4-jikkyou"
            skill.mkdir()
            (skill / "SKILL.md").write_bytes(b"a" * 101)
            with patch.object(mcp_skills, "ROOT", root), \
                 patch.object(mcp_skills, "MAX_SKILL_BYTES", 100):
                with self.assertRaisesRegex(ValueError, "size limit"):
                    mcp_skills.skill_text("jikkyou")

    def test_accepts_unicode_skill_at_exact_limit(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            skill = root / "ymm4-jikkyou"
            skill.mkdir()
            text = "魔理沙の実況"
            encoded = text.encode("utf-8")
            (skill / "SKILL.md").write_bytes(encoded)
            with patch.object(mcp_skills, "ROOT", root), \
                 patch.object(mcp_skills, "MAX_SKILL_BYTES", len(encoded)):
                self.assertEqual(mcp_skills.skill_text("jikkyou"), text)


if __name__ == "__main__":
    unittest.main()
