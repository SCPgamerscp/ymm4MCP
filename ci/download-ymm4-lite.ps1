# Download only the official YMM4 Lite reference archive. A truncated
# response must not produce a build or a published .ymme package.
$ErrorActionPreference = 'Stop'
$temporaryRoot = if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { [IO.Path]::GetTempPath() }
$archive = Join-Path $temporaryRoot 'YukkuriMovieMaker_v4_Lite.zip'
$directory = Join-Path $temporaryRoot 'ymm4-lite'
$source = 'https://files.manjubox.net/YukkuriMovieMaker_v4_Lite.zip'
$maxAttempts = 3

# Begin each invocation with a fresh archive. A failed HTTP transfer may leave
# usable partial bytes, but a completed/invalid ZIP must never be resumed.
Remove-Item -LiteralPath $archive -Force -ErrorAction SilentlyContinue

for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    $downloadCompleted = $false
    try {
        # A failed HTTP transfer can leave a partial ZIP. PowerShell's -Resume
        # uses HTTP Range where supported; if the server cannot resume, it
        # downloads from the beginning. Never reuse partially expanded DLLs.
        Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction SilentlyContinue
        Invoke-WebRequest -Uri $source -OutFile $archive -Resume -TimeoutSec 300 -ErrorAction Stop
        $downloadCompleted = $true
        if ((Get-Item -LiteralPath $archive).Length -lt 1MB) {
            throw 'Official YMM4 Lite download is unexpectedly small'
        }

        # Verify that the ZIP directory and the required entries are readable
        # before extracting. Expand-Archive will also detect corrupt members.
        $zip = [IO.Compression.ZipFile]::OpenRead($archive)
        try {
            $plugins = @($zip.Entries | Where-Object { $_.Name -ceq 'YukkuriMovieMaker.Plugin.dll' })
            $controls = @($zip.Entries | Where-Object { $_.Name -ceq 'YukkuriMovieMaker.Controls.dll' })
            if ($plugins.Count -ne 1 -or $controls.Count -ne 1) {
                throw 'Official YMM4 Lite archive has missing or duplicate reference DLLs'
            }
        }
        finally {
            $zip.Dispose()
        }

        Expand-Archive -LiteralPath $archive -DestinationPath $directory -ErrorAction Stop
        $pluginDlls = @(Get-ChildItem -LiteralPath $directory -Recurse -Filter 'YukkuriMovieMaker.Plugin.dll' -File)
        if ($pluginDlls.Count -ne 1) {
            throw "Expected one YMM4 Plugin DLL; found $($pluginDlls.Count)"
        }
        $libPath = $pluginDlls[0].DirectoryName
        if (-not (Test-Path -LiteralPath (Join-Path $libPath 'YukkuriMovieMaker.Controls.dll') -PathType Leaf)) {
            throw 'YMM4 Controls DLL is missing next to the Plugin DLL'
        }
        return $libPath
    }
    catch {
        # A completed but invalid ZIP is not a resumable transfer: restart it
        # from byte zero. An interrupted HTTP download may be resumed on retry.
        if ($downloadCompleted) {
            Remove-Item -LiteralPath $archive -Force -ErrorAction SilentlyContinue
        }
        if ($attempt -eq $maxAttempts) {
            throw "Official YMM4 Lite download/validation failed after $maxAttempts attempts: $($_.Exception.Message)"
        }
        Write-Warning "Official YMM4 Lite download/validation failed on attempt $attempt/$maxAttempts; retrying."
        Start-Sleep -Seconds (2 * $attempt)
    }
}
