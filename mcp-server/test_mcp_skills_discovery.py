"""MCP skill discovery and prompt input regression tests."""
import unittest
from unittest.mock import patch

import mcp_skills


class SkillDiscoveryTests(unittest.IsolatedAsyncioTestCase):
    async def test_pronunciation_skill_is_a_discoverable_resource(self):
        resources = await mcp_skills.list_resources()
        self.assertIn("ymm4://skills/chosei", [str(r.uri) for r in resources])
        with patch.object(mcp_skills, "skill_text", return_value="調声スキル") as read:
            self.assertEqual(await mcp_skills.read_resource("ymm4://skills/chosei"), "調声スキル")
            read.assert_called_once_with("chosei")

    async def test_pronunciation_resource_is_not_misrepresented_as_video_prompt(self):
        names = [p.name for p in await mcp_skills.list_prompts()]
        self.assertEqual(set(names), set(mcp_skills.SKILLS))
        self.assertNotIn("chosei", names)
        with self.assertRaisesRegex(ValueError, "Unknown production prompt"):
            await mcp_skills.get_prompt("chosei", {"theme": "こんにちは"})

    async def test_prompt_rejects_non_object_arguments(self):
        with patch.object(mcp_skills, "skill_text", return_value="テスト"):
            for args in ([], ["theme"], "theme", 1, True):
                with self.subTest(args=args), self.assertRaises(ValueError):
                    await mcp_skills.get_prompt("jikkyou", args)

    async def test_prompt_rejects_invalid_video_paths(self):
        with patch.object(mcp_skills, "skill_text", return_value="テスト"):
            for path in (None, 42, {}, "", "   ", "C:/clip\x00.mp4", "C:/clip\r.mp4", "C:/clip\x7f.mp4"):
                with self.subTest(path=path), self.assertRaisesRegex(ValueError, "video_path"):
                    await mcp_skills.get_prompt("jikkyou", {"theme": "実況", "video_path": path})

    async def test_prompt_keeps_existing_valid_inputs(self):
        with patch.object(mcp_skills, "skill_text", return_value="テスト"):
            result = await mcp_skills.get_prompt("jikkyou", {
                "theme": "ボス戦", "video_path": "C:/素材/clip.mp4", "duration_seconds": "60"})
        self.assertIn('"video_path": "C:/素材/clip.mp4"', result.messages[0].content.text)
        self.assertIn("plan_edit", result.messages[0].content.text)


if __name__ == "__main__":
    unittest.main()
