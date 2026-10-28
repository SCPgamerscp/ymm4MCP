# Exercise the real downloader with an interrupted transfer and a corrupt
# completed archive. No external YMM4 server or real YMM4 binaries are used.
$ErrorActionPreference = 'Stop'
$helper = Join-Path $PSScriptRoot 'download-ymm4-lite.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ("ymm4-lite-resume-test-" + [guid]::NewGuid().ToString('N'))
$originalRunnerTemp = $env:RUNNER_TEMP
New-Item -ItemType Directory -Path $testRoot | Out-Null

try {
    $fixture = Join-Path $testRoot 'fixture.zip'
    $zip = [IO.Compression.ZipFile]::Open($fixture, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in @('YukkuriMovieMaker.Plugin.dll', 'YukkuriMovieMaker.Controls.dll')) {
            $entry = $zip.CreateEntry($name)
            $stream = $entry.Open()
            try {
                $bytes = [Text.Encoding]::UTF8.GetBytes("fixture $name")
                $stream.Write($bytes, 0, $bytes.Length)
            }
            finally { $stream.Dispose() }
        }
        # The real helper rejects archives under 1 MB. Use an uncompressed
        # random payload to test the same validation path as the official ZIP.
        $entry = $zip.CreateEntry('filler.bin', [IO.Compression.CompressionLevel]::NoCompression)
        $stream = $entry.Open()
        try {
            $randomBytes = [byte[]]::new(1250000)
            [Security.Cryptography.RandomNumberGenerator]::Fill($randomBytes)
            $stream.Write($randomBytes, 0, $randomBytes.Length)
        }
        finally { $stream.Dispose() }
    }
    finally { $zip.Dispose() }

    $script:validZip = [IO.File]::ReadAllBytes($fixture)
    $half = [int][Math]::Floor($script:validZip.Length / 2)
    $script:partialZip = [byte[]]::new($half)
    [Array]::Copy($script:validZip, $script:partialZip, $half)
    $script:invalidZip = [byte[]]::new(1250000)
    [Security.Cryptography.RandomNumberGenerator]::Fill($script:invalidZip)
    $script:mode = ''
    $script:callCount = 0

    # Function shadowing deliberately replaces the HTTP transport only. The
    # actual helper's retry, ZIP verification and extraction run unchanged.
    function Invoke-WebRequest {
        [CmdletBinding()]
        param([string]$Uri, [string]$OutFile, [int]$TimeoutSec, [switch]$Resume)
        if ($Uri -ne 'https://files.manjubox.net/YukkuriMovieMaker_v4_Lite.zip') {
            throw 'Downloader changed its official source'
        }
        if (-not $Resume -or $TimeoutSec -ne 300) {
            throw 'Downloader must use bounded resumable requests'
        }
        $script:callCount++
        if ($script:callCount -eq 1 -and $script:mode -eq 'interrupted') {
            [IO.File]::WriteAllBytes($OutFile, $script:partialZip)
            throw 'Simulated ResponseEnded after partial transfer'
        }
        if ($script:callCount -eq 2 -and $script:mode -eq 'interrupted') {
            if (-not (Test-Path -LiteralPath $OutFile) -or
                (Get-Item -LiteralPath $OutFile).Length -ne $script:partialZip.Length) {
                throw 'Partial transfer was discarded instead of resumed'
            }
            $stream = [IO.File]::Open($OutFile, [IO.FileMode]::Append)
            try {
                $stream.Write($script:validZip, $script:partialZip.Length,
                    $script:validZip.Length - $script:partialZip.Length)
            }
            finally { $stream.Dispose() }
            return
        }
        if ($script:callCount -eq 1 -and $script:mode -eq 'corrupt') {
            [IO.File]::WriteAllBytes($OutFile, $script:invalidZip)
            return
        }
        if ($script:callCount -eq 2 -and $script:mode -eq 'corrupt') {
            if (Test-Path -LiteralPath $OutFile) {
                throw 'Completed invalid ZIP must be discarded before retry'
            }
            [IO.File]::WriteAllBytes($OutFile, $script:validZip)
            return
        }
        throw "Unexpected download attempt $script:callCount ($script:mode)"
    }

    foreach ($scenario in @('interrupted', 'corrupt')) {
        $env:RUNNER_TEMP = Join-Path $testRoot $scenario
        New-Item -ItemType Directory -Path $env:RUNNER_TEMP | Out-Null
        $script:mode = $scenario
        $script:callCount = 0
        $lib = & $helper
        if ($script:callCount -ne 2) {
            throw "$scenario expected exactly two download calls; got $script:callCount"
        }
        foreach ($name in @('YukkuriMovieMaker.Plugin.dll', 'YukkuriMovieMaker.Controls.dll')) {
            if (-not (Test-Path -LiteralPath (Join-Path $lib $name) -PathType Leaf)) {
                throw "$scenario failed to extract $name"
            }
        }
        Write-Output "PASS: $scenario"
    }
}
finally {
    $env:RUNNER_TEMP = $originalRunnerTemp
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
