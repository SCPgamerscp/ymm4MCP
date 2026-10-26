"""Contract tests for the guarded YMM4 Release workflow (Issue #22).

The release job cannot be executed by a pull-request event. These tests
prevent regressions in its source-level safety gates; an actual publish still
requires a maintainer-triggered workflow run on main.
"""
from pathlib import Path
import unittest


WORKFLOW = (Path(__file__).resolve().parents[1]
            / ".github" / "workflows" / "package-ymme.yml")


class ReleaseWorkflowContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.workflow = WORKFLOW.read_text(encoding="utf-8-sig")
        cls.release = cls.workflow.split("\n  release:\n", 1)[1]

    def test_release_requires_both_test_and_package_jobs(self):
        self.assertIn("needs: [package, python-tests]", self.release)
        self.assertIn("group: ymm4mcp-release", self.release)
        self.assertIn("cancel-in-progress: false", self.release)

    def test_manual_release_rejects_stale_main_and_existing_tags(self):
        self.assertIn("$env:GITHUB_REF -ne 'refs/heads/main'", self.release)
        self.assertIn("$current -ne $latest", self.release)
        self.assertIn("gh release view $tag --json tagName", self.release)
        self.assertIn('gh api "repos/$env:GITHUB_REPOSITORY/git/ref/tags/$tag"', self.release)

    def test_tag_rerun_verifies_existing_release_without_overwrite(self):
        self.assertIn("Release $tag already exists; verifying without modifying it",
                      self.release)
        self.assertIn("gh release download $tag --dir $verifyDir --clobber",
                      self.release)
        self.assertIn("Compare-Object $expectedAssets $publishedAssets",
                      self.release)
        self.assertIn("if ($createdRelease)", self.release)
        self.assertIn("Get-FileHash -LiteralPath $downloaded -Algorithm SHA256",
                      self.release)
        self.assertIn("[IO.Compression.ZipFile]::OpenRead($downloaded)",
                      self.release)


if __name__ == "__main__":
    unittest.main()
