"""Regression contract for official YMM4 reference downloads in package CI.

Windows executes the PowerShell helper in Actions; this test catches accidental
reversion to a one-shot unvalidated download on every Python test run.
"""
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class Ymm4ReferenceDownloadContractTests(unittest.TestCase):
    def test_package_uses_bounded_validated_download(self):
        script = (ROOT / "ci" / "download-ymm4-lite.ps1").read_text(encoding="utf-8")
        workflow = (ROOT / ".github" / "workflows" / "package-ymme.yml").read_text(encoding="utf-8")
        self.assertIn("$maxAttempts = 3", script)
        self.assertIn("Invoke-WebRequest -Uri $source", script)
        self.assertIn("-TimeoutSec 300", script)
        self.assertIn("Remove-Item -LiteralPath $archive", script)
        self.assertIn("Remove-Item -LiteralPath $directory", script)
        self.assertIn("[IO.Compression.ZipFile]::OpenRead($archive)", script)
        self.assertIn("Expand-Archive -LiteralPath $archive", script)
        self.assertIn("YukkuriMovieMaker.Plugin.dll", script)
        self.assertIn("YukkuriMovieMaker.Controls.dll", script)
        self.assertIn("throw \"Official YMM4 Lite download/validation failed", script)
        self.assertIn("& ./ci/download-ymm4-lite.ps1", workflow)
        self.assertNotIn("Invoke-WebRequest -Uri 'https://files.manjubox.net/", workflow)


if __name__ == "__main__":
    unittest.main()
