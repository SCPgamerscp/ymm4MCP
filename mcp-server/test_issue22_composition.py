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


if __name__ == "__main__":
    unittest.main()
