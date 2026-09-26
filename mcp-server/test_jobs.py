"""Export/job request validation and MCP dispatch. No YMM4 required."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

import editing
import jobs
import server


class ExportValidationTests(unittest.TestCase):
    def test_absolute_path_and_format_required(self):
        parsed = jobs.validate_export_request({"path": "C:/Videos/out.mp4"})
        self.assertEqual(parsed["format"], "mp4")
        self.assertFalse(parsed["overwrite"])
        self.assertEqual(parsed["timeout_seconds"], 1800)
        self.assertEqual(jobs.validate_export_request({"output_path": "/tmp/a.wav", "format": "wav"})["format"], "wav")

    def test_rejects_relative_and_invalid_values(self):
        invalid = [
            {},
            {"path": "out.mp4"},
            {"path": "C:/out.mp4", "format": "exe"},
            {"path": "C:/out.mp4", "format": "wav"},
            {"path": "C:/out.mp4", "overwrite": "true"},
            {"path": "C:/out.mp4", "timeout_seconds": 0},
            {"path": "C:/out.mp4", "timeout_seconds": True},
            {"path": "C:/out.mp4", "idempotency_key": " x"},
            {"path": "http://127.0.0.1/out.mp4"},
            {"path": ""},
        ]
        for args in invalid:
            with self.subTest(args=args), self.assertRaises(ValueError):
                jobs.validate_export_request(args)

    def test_project_path_must_be_ymmp(self):
        self.assertEqual(jobs.validate_project_path({"path": "C:/p/a.ymmp"}), "C:/p/a.ymmp")
        for args in ({}, {"path": "C:/p/a.txt"}, {"path": "rel.ymmp"}):
            with self.subTest(args=args), self.assertRaises(ValueError):
                jobs.validate_project_path(args)

    def test_job_id_pattern(self):
        self.assertTrue(jobs.job_id_ok("job_" + "a" * 32))
        self.assertFalse(jobs.job_id_ok("job_"))
        self.assertFalse(jobs.job_id_ok("../jobs"))
        self.assertFalse(jobs.job_id_ok("job_abc/cancel"))


class MediaVerifyTests(unittest.TestCase):
    def test_mp4_ftyp_and_wav_riff(self):
        with tempfile.TemporaryDirectory() as directory:
            mp4 = Path(directory) / "clip.mp4"
            mp4.write_bytes(b"\x00\x00\x00\x20ftypisom" + b"\x00" * 24)
            wav = Path(directory) / "audio.wav"
            wav.write_bytes(b"RIFF\x00\x00\x00\x00WAVE" + b"\x00" * 24)
            empty = Path(directory) / "empty.mp4"
            empty.write_bytes(b"not-an-mp4-header!!!!" + b"\x00" * 20)
            self.assertTrue(jobs.verify_media_file(str(mp4), "mp4")["ok"])
            self.assertTrue(jobs.verify_media_file(str(wav), "wav")["ok"])
            self.assertEqual(jobs.verify_media_file(str(empty), "mp4")["error_code"], "EXPORT_VERIFY_FAILED")
            missing = jobs.verify_media_file(str(Path(directory) / "gone.mp4"), "mp4")
            self.assertEqual(missing["error_code"], "EXPORT_FILE_MISSING")
            tiny = Path(directory) / "tiny.mp4"
            tiny.write_bytes(b"short")
            self.assertEqual(jobs.verify_media_file(str(tiny), "mp4")["error_code"], "EXPORT_FILE_EMPTY")

    def test_unc_and_newline_paths(self):
        parsed = jobs.validate_export_request({"path": r"\\nas\share\out.mp4"})
        self.assertEqual(parsed["format"], "mp4")
        with self.assertRaises(ValueError):
            jobs.validate_export_request({"path": "C:/out\n.mp4"})


class JobDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_export_is_validated_before_http(self):
        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "export", "path": "out.mp4"})
            post.assert_not_awaited()
            await server.dispatch({
                "action": "control", "sub_action": "export", "path": "C:/Videos/final.mp4",
                "overwrite": True, "idempotency_key": "run-1", "timeout_seconds": 60,
            })
            post.assert_awaited_once_with("/project/export", {
                "path": "C:/Videos/final.mp4", "format": "mp4", "overwrite": True,
                "timeout_seconds": 60, "idempotency_key": "run-1",
            })

    async def test_job_poll_cancel_resume_and_open(self):
        job_id = "job_" + "b" * 32
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "jobs": []})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "jobs"})
            get.assert_awaited_once_with("/jobs")
        with patch.object(server, "ymm4_get", AsyncMock(return_value={"success": True, "job_id": job_id})) as get:
            await server.dispatch({"action": "get_info", "sub_action": "job", "job_id": job_id})
            get.assert_awaited_once_with(f"/jobs/{job_id}")
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            await server.dispatch({"action": "control", "sub_action": "cancel_job", "job_id": job_id})
            post.assert_awaited_once_with(f"/jobs/{job_id}/cancel")
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            await server.dispatch({"action": "control", "sub_action": "resume_job", "job_id": job_id})
            post.assert_awaited_once_with(f"/jobs/{job_id}/resume")
        with patch.object(server, "ymm4_post", AsyncMock(return_value={"success": True})) as post:
            await server.dispatch({"action": "control", "sub_action": "open", "path": "C:/p/a.ymmp"})
            post.assert_awaited_once_with("/project/open", {"path": "C:/p/a.ymmp"})
            post.reset_mock()
            await server.dispatch({"action": "control", "sub_action": "open", "path": "C:/p/a.ymmp", "force": True})
            post.assert_awaited_once_with("/project/open", {"path": "C:/p/a.ymmp", "force": True})
            post.reset_mock()
            await server.dispatch({"action": "control", "sub_action": "save_as", "path": "C:/p/b.ymmp", "overwrite": True})
            post.assert_awaited_once_with("/project/save-as", {"path": "C:/p/b.ymmp", "overwrite": True})

        with patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "save_as", "path": "C:/p/b.ymmp", "overwrite": "true"})
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "open", "path": "C:/p/a.ymmp", "force": "true"})
            post.assert_not_awaited()
            with self.assertRaises(ValueError):
                await server.dispatch({"action": "control", "sub_action": "export", "path": "C:/Videos/final.mp4", "timeout_seconds": "60"})
            post.assert_not_awaited()

    async def test_invalid_job_id_never_touches_http(self):
        with patch.object(server, "ymm4_get", new_callable=AsyncMock) as get, patch.object(server, "ymm4_post", new_callable=AsyncMock) as post:
            for sub in ("job",):
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "get_info", "sub_action": sub, "job_id": "../x"})
            for sub in ("cancel_job", "resume_job"):
                with self.assertRaises(ValueError):
                    await server.dispatch({"action": "control", "sub_action": sub, "job_id": "job_ab"})
            get.assert_not_awaited()
            post.assert_not_awaited()

    async def test_skill_prompt_mentions_export_jobs(self):
        import mcp_skills
        prompt = await mcp_skills.get_prompt("jikkyou", {"theme": "ボス戦"})
        self.assertIn("control/export", prompt.messages[0].content.text)
        self.assertIn("get_info/job", prompt.messages[0].content.text)


class FiniteHelperStillWorks(unittest.TestCase):
    def test_editing_helpers_unchanged(self):
        self.assertEqual(editing.finite_number("2", "n"), 2.0)


if __name__ == "__main__":
    unittest.main()
