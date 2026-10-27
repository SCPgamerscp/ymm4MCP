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

    def test_manual_release_requires_latest_main_and_never_overwrites_release(self):
        self.assertIn("$env:GITHUB_REF -ne 'refs/heads/main'", self.release)
        self.assertIn("$current -ne $latest", self.release)
        self.assertIn("gh release view $tag --json tagName", self.release)
        self.assertIn("Release $tag already exists; bump Version", self.release)

    def test_manual_release_recovers_only_tag_on_same_tested_commit(self):
        self.assertIn("git ls-remote --exit-code --refs origin $tagRef", self.release)
        self.assertIn("$tagProbeExit -eq 2", self.release)
        self.assertIn("$tagProbeExit -ne 0", self.release)
        self.assertIn('gh api -X POST "repos/$env:GITHUB_REPOSITORY/git/refs"', self.release)
        self.assertIn('-f "ref=$tagRef" -f "sha=$env:GITHUB_SHA"', self.release)
        self.assertIn("git fetch --no-tags origin $tagRef", self.release)
        self.assertIn("git rev-list -n 1 FETCH_HEAD", self.release)
        self.assertIn("$tagCommit -ne $env:GITHUB_SHA", self.release)
        self.assertIn("Tag $tag points to a different commit", self.release)
        self.assertIn("gh release create $tag $packagePath $serverPackage[0].FullName --verify-tag", self.release)
        self.assertNotIn("gh release create $tag $packagePath $serverPackage[0].FullName --target", self.release)
        self.assertIn("$createdRelease = $true", self.release)

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
