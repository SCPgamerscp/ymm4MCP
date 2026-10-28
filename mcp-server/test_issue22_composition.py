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
        # This order must match the numeric all-open-PR merge audit.
        self.assertIn(
            "for number in 218 219 221 228 229 234 241 243 244 248 251 254 255 257 259;",
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
        self.assertIn("needs: [setup-cross-pr-integration, all-open-prs-windows-package]", self.source)
        self.assertIn("tar -xf issue22-composed.tar -C composed", self.source)
        self.assertIn("python -m unittest discover -s mcp-server -p 'test_*.py' -q",
                      self.source)
        self.assertIn(
            "python -m unittest -v test_decoded_media_real_ffmpeg test_decoded_media_ffmpeg_integration test_decoded_media_required_tail_20261028",
            self.source,
        )


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


    def test_curated_suite_does_not_repeat_full_windows_package_download(self):
        # The 15-PR regression suite previously repeated the large YMM4 Lite
        # download and timed out while the full 50-PR package had passed.
        # One authoritative Windows package job must gate real export QA.
        curated = self.source.split(
            "  setup-cross-pr-integration:", 1
        )[1].split("  composed-real-export-qa:", 1)[0]
        all_pr_windows = self.source.split(
            "  all-open-prs-windows-package:", 1
        )[1]
        self.assertNotIn("Download official YMM4 Lite references", curated)
        self.assertNotIn("Build and verify composed YMM4 .ymme", curated)
        self.assertIn(
            "needs: [setup-cross-pr-integration, all-open-prs-windows-package]",
            self.source,
        )
        self.assertIn("Build complete all-PR .ymme package", all_pr_windows)
        self.assertEqual(
            self.source.count("      - name: Download official YMM4 Lite references"),
            1,
        )

    def test_all_open_prs_are_tested_from_one_immutable_archive(self):
        # A clean 50-PR merge is not enough: both platforms must consume the
        # same archive and run real regressions/builds without remote writes.
        self.assertIn("conflicts: ${{ steps.merge-audit.outputs.conflicts }}", self.source)
        self.assertIn("id: merge-audit", self.source)
        self.assertIn('echo "conflicts=$failed" >> "$GITHUB_OUTPUT"', self.source)
        self.assertIn("git archive --format=tar --output=all-open-prs-composed.tar HEAD", self.source)
        self.assertIn("name: all-open-prs-composed-source", self.source)
        for job in ("all-open-prs-python-integration", "all-open-prs-windows-package"):
            self.assertIn(f"  {job}:", self.source)
            self.assertIn("needs.audit-open-pr-merge-compatibility.outputs.conflicts == '0'",
                          self.source)
        self.assertIn("tar -xf all-open-prs-composed.tar -C composed", self.source)
        self.assertIn("python -m unittest discover -s mcp-server -p 'test_*.py' -q",
                      self.source)
        self.assertIn("dotnet build composed/YMM4McpPlugin/YMM4McpPlugin.csproj",
                      self.source)
        self.assertIn("Get-ChildItem -Path composed/tests -Recurse -Filter 'Keyframe*.csproj'",
                      self.source)
        self.assertNotIn("git push", self.source)

    def test_all_open_pr_merge_audit_is_read_only_and_complete(self):
        # Individual mergeable=true flags do not prove a 50-PR sequence works.
        # The advisory job discovers current open PRs and attempts sequential
        # local merges, recording conflicts without pushing or masking them.
        self.assertIn("  audit-open-pr-merge-compatibility:", self.source)
        self.assertIn("  pull-requests: read", self.source)
        self.assertIn(
            'gh api --paginate "repos/$GITHUB_REPOSITORY/pulls?state=open&per_page=100"',
            self.source,
        )
        self.assertIn('mapfile -t numbers < <(sort -n -u', self.source)
        self.assertIn('git checkout -B audit-open-prs refs/remotes/origin/main', self.source)
        self.assertIn('git merge --no-ff --no-edit "refs/remotes/audit-pr/${number}"', self.source)
        self.assertIn('git merge --abort', self.source)
        self.assertIn('echo "**Composed: ', self.source)
        self.assertIn('persist-credentials: false', self.source)
        self.assertNotIn('git push', self.source)


if __name__ == "__main__":
    unittest.main()
