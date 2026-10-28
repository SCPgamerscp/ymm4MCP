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
            [System.Security.Cryptography.RandomNumberGenerator]::Fill($randomBytes)
            $stream.Write($randomBytes, 0, $randomBytes.Length)
        }
        finally { $stream.Dispose() }
    }
    finally { $zip.Dispose() }

    $global:validZip = [IO.File]::ReadAllBytes($fixture)
    $half = [int][Math]::Floor($global:validZip.Length / 2)
    $global:partialZip = [byte[]]::new($half)
    [Array]::Copy($global:validZip, $global:partialZip, $half)
    $global:invalidZip = [byte[]]::new(1250000)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($global:invalidZip)
    $global:mode = ''
    $global:callCount = 0

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
        $global:callCount++
        if ($global:callCount -eq 1 -and $global:mode -eq 'interrupted') {
            [IO.File]::WriteAllBytes($OutFile, $global:partialZip)
            throw 'Simulated ResponseEnded after partial transfer'
        }
        if ($global:callCount -eq 2 -and $global:mode -eq 'interrupted') {
            if (-not (Test-Path -LiteralPath $OutFile) -or
                (Get-Item -LiteralPath $OutFile).Length -ne $global:partialZip.Length) {
                throw 'Partial transfer was discarded instead of resumed'
            }
            $stream = [IO.File]::Open($OutFile, [IO.FileMode]::Append)
            try {
                $stream.Write($global:validZip, $global:partialZip.Length,
                    $global:validZip.Length - $global:partialZip.Length)
            }
            finally { $stream.Dispose() }
            return
        }
        if ($global:callCount -eq 1 -and $global:mode -eq 'corrupt') {
            [IO.File]::WriteAllBytes($OutFile, $global:invalidZip)
            return
        }
        if ($global:callCount -eq 2 -and $global:mode -eq 'corrupt') {
            if (Test-Path -LiteralPath $OutFile) {
                throw 'Completed invalid ZIP must be discarded before retry'
            }
            [IO.File]::WriteAllBytes($OutFile, $global:validZip)
            return
        }
        throw "Unexpected download attempt $global:callCount ($global:mode)"
    }

    foreach ($scenario in @('interrupted', 'corrupt')) {
        $env:RUNNER_TEMP = Join-Path $testRoot $scenario
        New-Item -ItemType Directory -Path $env:RUNNER_TEMP | Out-Null
        $global:mode = $scenario
        $global:callCount = 0
        $lib = & $helper
        if ($global:callCount -ne 2) {
            throw "$scenario expected exactly two download calls; got $global:callCount"
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
    Remove-Variable -Name validZip,partialZip,invalidZip,mode,callCount -Scope Global -ErrorAction SilentlyContinue
}
