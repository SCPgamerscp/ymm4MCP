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

    def test_server_archive_excludes_untracked_credentials(self):
        # This regression protects the published ZIP, not only the installer.
        self.assertIn('git archive --format=zip "--output=$serverPackage" HEAD mcp-server',
                      self.workflow)
        self.assertIn('[IO.File]::WriteAllText($localToken', self.workflow)
        self.assertIn('[IO.File]::WriteAllText($localEnv', self.workflow)
        self.assertIn('Sensitive or ephemeral files in Python release ZIP',
                      self.workflow)
        self.assertNotIn('Compress-Archive -Path mcp-server', self.workflow)

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

    def test_tag_event_must_match_built_package_version_before_publish(self):
        # A tag such as v999 pointing to main must not publish v1.2.3
        # binaries as v999 merely because the tag is an ancestor of main.
        self.assertIn("$env:GITHUB_REF_NAME -cne $tag", self.release)
        self.assertIn("does not match packaged version", self.release)
        self.assertLess(
            self.release.index("$env:GITHUB_REF_NAME -cne $tag"),
            self.release.index("gh release create $tag $packagePath"),
        )
        self.assertLess(
            self.release.index("$env:GITHUB_REF_NAME -cne $tag"),
            self.release.index("gh release create $env:GITHUB_REF_NAME"),
        )

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


    def test_official_reference_download_retries_and_validates_archive(self):
        # Transient upstream truncation must not fail the first attempt forever,
        # and no partial ZIP/DLL may be used for a successful package.
        root = WORKFLOW.parents[2]
        script = (root / "ci" / "download-ymm4-lite.ps1").read_text(encoding="utf-8")
        composed = (root / ".github" / "workflows"
                    / "issue22-cross-pr-integration.yml").read_text(encoding="utf-8")
        self.assertIn("$maxAttempts = 3", script)
        self.assertIn("-TimeoutSec 300", script)
        self.assertIn("Remove-Item -LiteralPath $archive", script)
        self.assertIn("Remove-Item -LiteralPath $directory", script)
        self.assertIn("[IO.Compression.ZipFile]::OpenRead($archive)", script)
        self.assertIn("Expand-Archive -LiteralPath $archive", script)
        self.assertIn("YukkuriMovieMaker.Plugin.dll", script)
        self.assertIn("YukkuriMovieMaker.Controls.dll", script)
        self.assertIn("throw \"Official YMM4 Lite download/validation failed", script)
        self.assertIn("& ./ci/download-ymm4-lite.ps1", self.workflow)
        self.assertIn("& ./ci/download-ymm4-lite.ps1", composed)
        self.assertIn("& ./composed/ci/download-ymm4-lite.ps1", composed)
        self.assertNotIn("Invoke-WebRequest -Uri 'https://files.manjubox.net/", self.workflow)
        self.assertNotIn("Invoke-WebRequest -Uri 'https://files.manjubox.net/", composed)


if __name__ == "__main__":
    unittest.main()
