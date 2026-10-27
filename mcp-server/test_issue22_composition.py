"""Regression contract for Issue #22/#140 composed PR CI.

A queued pull_request check must test the #228 head SHA from its event, not a
newer revision fetched from the moving refs/pull/228/head. Manual dispatches
are allowed to compose the latest available heads.
"""
from pathlib import Path
import unittest


WORKFLOW = (
    Path(__file__).resolve().parents[1]
    / ".github" / "workflows" / "issue22-cross-pr-integration.yml"
)


class ComposedPullRequestHeadTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = WORKFLOW.read_text(encoding="utf-8-sig")

    def test_pull_request_uses_triggering_head_sha(self):
        self.assertIn(
            "EVENT_PR_228_HEAD_SHA: ${{ github.event.pull_request.head.sha || '' }}",
            self.source,
        )
        self.assertIn(
            'if [[ "$number" == 228 && -n "$EVENT_PR_228_HEAD_SHA" ]]; then',
            self.source,
        )
        self.assertIn(
            'git fetch --no-tags origin "$EVENT_PR_228_HEAD_SHA:refs/remotes/pr/${number}"',
            self.source,
        )

    def test_manual_dispatch_keeps_current_ref_fallback(self):
        self.assertIn("workflow_dispatch:", self.source)
        self.assertIn(
            'git fetch --no-tags origin "refs/pull/${number}/head:refs/remotes/pr/${number}"',
            self.source,
        )
        self.assertIn('git merge --no-ff --no-edit "refs/remotes/pr/${number}"', self.source)


    def test_composition_includes_export_qa_pr(self):
        # The export QA branch modifies the same packaging workflow as #228;
        # composing it is required before claiming release integration passed.
        self.assertIn(
            "for number in 221 228 229 234 241 243 244 248 257 259 218 219 251 254 255;",
            self.source,
        )

    def test_export_qa_discovers_new_regression_modules(self):
        # Hard-coded unittest module lists missed newly added tests on PR #229,
        # giving a green cross-PR build without testing the newest QA changes.
        self.assertIn("Get-ChildItem -LiteralPath . -Filter 'test_decoded_media_*.py' -File", self.source)
        self.assertIn("if ($qaTests.Count -lt 8)", self.source)
        self.assertIn(
            "python -m unittest discover -s . -p 'test_decoded_media_*.py' -q",
            self.source,
        )
        self.assertNotIn(
            "python -m unittest -q test_decoded_media_qa test_decoded_media_tail_and_output",
            self.source,
        )


    def test_real_export_qa_uses_exact_composed_source(self):
        # A second merge on Ubuntu could pick up newer PR heads than the
        # Windows build. Transfer the built composition as a source artifact.
        self.assertIn("git archive --format=tar --output=issue22-composed.tar HEAD",
                      self.source)
        self.assertIn("name: issue22-composed-source", self.source)
        self.assertIn("  composed-real-export-qa:", self.source)
        self.assertIn("needs: setup-cross-pr-integration", self.source)
        self.assertIn("tar -xf issue22-composed.tar -C composed", self.source)
        self.assertIn("python -m unittest discover -s mcp-server -p 'test_*.py' -q",
                      self.source)
        self.assertIn("python -m unittest -v test_decoded_media_real_ffmpeg",
                      self.source)


    def test_real_export_qa_uses_runner_ffmpeg_before_apt(self):
        # Unconditional apt update/install can stall the cross-platform
        # verification even when both binaries are already on the runner.
        self.assertIn(
            "if command -v ffmpeg >/dev/null 2>&1 && command -v ffprobe >/dev/null 2>&1; then",
            self.source,
        )
        self.assertIn("sudo apt-get install -y -qq ffmpeg", self.source)
        self.assertIn("ffmpeg -version | head -n 1", self.source)
        self.assertIn("ffprobe -version | head -n 1", self.source)


if __name__ == "__main__":
    unittest.main()
