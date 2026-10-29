param([switch]$ConfigureClaudeDesktop, [switch]$CheckConnection,
      [ValidateRange(1024, 65535)][Nullable[int]]$Port, [switch]$Install)
$ErrorActionPreference = 'Stop'
$sourceDir = $PSScriptRoot
$serverDir = $sourceDir

# The Windows py launcher can exist without a registered Python 3 runtime.
# Probe it first, then fall back to python.exe instead of failing installation.
# The MCP dependency stack requires Python >= 3.10.  A machine may have an
# older py -3 default even when python.exe points to a supported interpreter.
$pythonCommand = $null
$pythonArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) {
    # A user-configured py -3 default may be Python 3.9 even when a newer
    # interpreter is registered. -V:3 asks the modern Windows launcher for
    # the newest registered Python 3 (including versions after 3.14). Older
    # launchers can reject it, so retain the explicit-version fallbacks.
    foreach ($selector in @('-3', '-V:3', '-3.14', '-3.13', '-3.12', '-3.11', '-3.10')) {
        try {
            & py $selector -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>$null
            if ($LASTEXITCODE -eq 0) {
                $pythonCommand = 'py'
                $pythonArgs = @($selector)
                break
            }
        } catch {
            # A missing launcher registration is not fatal: try the next one.
        }
    }
}
if ($null -eq $pythonCommand -and (Get-Command python -ErrorAction SilentlyContinue)) {
    try {
        & python -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>$null
        if ($LASTEXITCODE -eq 0) { $pythonCommand = 'python' }
    } catch {
        # Report a single actionable error below if both candidates fail.
    }
}
if ($null -eq $pythonCommand) {
    throw '利用可能な Python 3.10 以上が見つかりません。Python 3.10 以上をインストールしてから再実行してください。'
}
# Install into a persistent per-user directory.  Running from a downloaded ZIP
# extraction directory would break Claude Desktop when that folder is removed.
# Configuring Claude Desktop must always use the persistent installation:
# otherwise a later deletion of the downloaded ZIP folder breaks the MCP entry.
# Read-only connection checks without -Install remain in-place.
if ($Install -or $ConfigureClaudeDesktop) {
    if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA が見つかりません。' }
    $targetDir = Join-Path $env:LOCALAPPDATA 'YMM4MCP\mcp-server'
    if ([IO.Path]::GetFullPath($sourceDir).TrimEnd('\') -ine
        [IO.Path]::GetFullPath($targetDir).TrimEnd('\')) {
        [IO.Directory]::CreateDirectory($targetDir) | Out-Null
        # Keep local MCP configuration across upgrades.  A Release ZIP may
        # contain a default config.json, but must not erase an existing user's
        # model selection or local secrets.  Never import connection.json from
        # the extracted ZIP: the plugin owns the live connection descriptor.
        # New installs still receive the packaged default config.json.
        $copyArgs = @('/E', '/XJ', '/XD', '.venv', '__pycache__', '.git',
                      '/XF', '*.pyc', 'connection.json')
        foreach ($localName in @('config.json', 'settings.local.json', '.env')) {
            if (Test-Path -LiteralPath (Join-Path $targetDir $localName) -PathType Leaf) {
                $copyArgs += $localName
            }
        }
        $copyArgs += @('/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
        # robocopy preserves the directory layout and skips ephemeral Python
        # environments/caches. Exit codes 0..7 mean success on Windows.
        & robocopy.exe $sourceDir $targetDir @copyArgs
        if ($LASTEXITCODE -ge 8) { throw 'MCP サーバーの永続インストールに失敗しました。' }
    }
    $serverDir = $targetDir
}
$venvDir = Join-Path $serverDir '.venv'
$venvPython = Join-Path $venvDir 'Scripts\python.exe'
& $pythonCommand @pythonArgs -m venv $venvDir
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython)) {
    throw 'Python の仮想環境を作成できませんでした。'
}

& $venvPython -m pip install --disable-pip-version-check -r (Join-Path $serverDir 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'MCP サーバーの依存関係をインストールできませんでした。' }

$entry = @{ command = $venvPython; args = @((Join-Path $serverDir 'server.py')) }
if ($null -ne $Port) {
    $entry.env = @{ YMM4_API_BASE = "http://127.0.0.1:$Port/api" }
}
# Validate connectivity before changing Claude Desktop settings. A failed
# -CheckConnection must not leave a broken mcpServers.ymm4 entry behind.
if ($CheckConnection) {
    $oldBase = $env:YMM4_API_BASE
    try {
        if ($null -ne $Port) { $env:YMM4_API_BASE = $entry.env.YMM4_API_BASE }
        & $venvPython (Join-Path $serverDir 'check_connection.py')
        if ($LASTEXITCODE -ne 0) { throw 'YMM4プラグインへの接続確認に失敗しました。' }
    } finally {
        $env:YMM4_API_BASE = $oldBase
    }
}

if ($ConfigureClaudeDesktop) {
    if (-not $env:APPDATA) { throw 'APPDATA が見つかりません。Claude Desktop の設定場所を確認してください。' }
    $configPaths = @((Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'))
    # Configure both classic and initialized Microsoft Store installations.
    if ($env:LOCALAPPDATA) {
        $packagesDir = Join-Path $env:LOCALAPPDATA 'Packages'
        if (Test-Path -LiteralPath $packagesDir) {
            foreach ($package in @(Get-ChildItem -LiteralPath $packagesDir -Directory -Filter 'Claude_*')) {
                $storeConfigDir = Join-Path $package.FullName 'LocalCache\Roaming\Claude'
                if (Test-Path -LiteralPath $storeConfigDir) {
                    $configPaths += (Join-Path $storeConfigDir 'claude_desktop_config.json')
                }
            }
        }
    }

    # Stage ALL new configs before touching any original. A write failure on a
    # Store install must not silently leave the classic install half-updated.
    $updates = New-Object System.Collections.ArrayList
    $committed = New-Object System.Collections.ArrayList
    try {
        foreach ($configPath in $configPaths) {
            $exists = Test-Path -LiteralPath $configPath
            if ($exists) {
                $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($null -eq $config -or $config -isnot [pscustomobject] -or
                    ($null -ne $config.mcpServers -and $config.mcpServers -isnot [pscustomobject])) {
                    throw "Claude Desktop の設定は JSON オブジェクトである必要があります: $configPath"
                }
            } else {
                $config = [pscustomobject]@{}
            }
            if ($null -eq $config.mcpServers) {
                $config | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{}) -Force
            }
            $current = $config.mcpServers.ymm4
            if ($null -ne $current -and $current -isnot [pscustomobject]) {
                throw "既存の mcpServers.ymm4 は JSON オブジェクトである必要があります: $configPath"
            }
            # Keep user-managed MCP settings (for example API credentials,
            # timeout and disabled flags) when reinstalling or upgrading.
            # Only the executable/script paths and an explicitly requested
            # -Port belong to the installer. Do not mutate $current.env:
            # the original config must remain intact until the atomic commit.
            $desiredEntry = [pscustomobject]@{
                command = $entry.command
                args = @($entry.args)
            }
            if ($null -ne $current) {
                foreach ($property in $current.PSObject.Properties) {
                    if ($property.Name -notin @('command', 'args', 'env')) {
                        $desiredEntry | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
                    }
                }
            }
            $desiredEnv = [pscustomobject]@{}
            if ($null -ne $current -and $null -ne $current.env) {
                if ($current.env -isnot [pscustomobject]) {
                    throw "既存の mcpServers.ymm4.env は JSON オブジェクトである必要があります: $configPath"
                }
                foreach ($property in $current.env.PSObject.Properties) {
                    $desiredEnv | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
                }
            }
            if ($null -ne $Port) {
                $desiredEnv | Add-Member -NotePropertyName YMM4_API_BASE -NotePropertyValue $entry.env.YMM4_API_BASE -Force
            }
            if (@($desiredEnv.PSObject.Properties).Count -gt 0) {
                $desiredEntry | Add-Member -NotePropertyName env -NotePropertyValue $desiredEnv
            }
            $sameEntry = $null -ne $current -and
                         $current.command -ceq $entry.command -and
                         ($current.args -is [array]) -and
                         @($current.args).Count -eq 1 -and
                         @($current.args)[0] -ceq $entry.args[0] -and
                         ($null -eq $Port -or
                          ($null -ne $current.env -and
                           $current.env.YMM4_API_BASE -ceq $entry.env.YMM4_API_BASE))
            if ($sameEntry) {
                Write-Output "Claude Desktop の設定は更新済みです: $configPath"
                continue
            }

            $config.mcpServers | Add-Member -NotePropertyName ymm4 -NotePropertyValue $desiredEntry -Force
            $configDir = Split-Path -Parent $configPath
            [IO.Directory]::CreateDirectory($configDir) | Out-Null
            $tempPath = "$configPath.tmp.$([guid]::NewGuid().ToString('N'))"
            $backupPath = if ($exists) { "$configPath.bak.$([guid]::NewGuid().ToString('N'))" } else { $null }
            $update = [pscustomobject]@{
                Path = $configPath
                Temp = $tempPath
                Backup = $backupPath
                Existed = $exists
            }
            [void]$updates.Add($update)
            $json = $config | ConvertTo-Json -Depth 100
            [IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
        }

        try {
            foreach ($update in $updates) {
                if ($update.Existed) {
                    # Keep a byte-for-byte backup of each original configuration.
                    [IO.File]::Replace($update.Temp, $update.Path, $update.Backup)
                } else {
                    [IO.File]::Move($update.Temp, $update.Path)
                }
                [void]$committed.Add($update)

                # CI-only fault injection verifies rollback across both configs.
                if ($env:CI -eq 'true' -and
                    $env:YMM4MCP_TEST_FAIL_AFTER_FIRST_CONFIG_WRITE -eq '1' -and
                    $committed.Count -eq 1) {
                    throw 'Simulated failure after first Claude config write'
                }
            }
        } catch {
            $originalError = $_
            $rollbackErrors = New-Object System.Collections.ArrayList
            for ($i = $committed.Count - 1; $i -ge 0; $i--) {
                $update = $committed[$i]
                try {
                    if ($update.Existed) {
                        $restoreTemp = "$($update.Path).restore.$([guid]::NewGuid().ToString('N'))"
                        try {
                            [IO.File]::Copy($update.Backup, $restoreTemp)
                            $restoreBackup = "$($update.Path).rollback.$([guid]::NewGuid().ToString('N'))"
                            [IO.File]::Replace($restoreTemp, $update.Path, $restoreBackup)
                            Remove-Item -LiteralPath $restoreBackup -ErrorAction SilentlyContinue
                        } finally {
                            if (Test-Path -LiteralPath $restoreTemp) {
                                Remove-Item -LiteralPath $restoreTemp -ErrorAction SilentlyContinue
                            }
                        }
                    } else {
                        Remove-Item -LiteralPath $update.Path -Force -ErrorAction Stop
                    }
                } catch {
                    [void]$rollbackErrors.Add("$($update.Path): $($_.Exception.Message); backup=$($update.Backup)")
                }
            }
            if ($rollbackErrors.Count -gt 0) {
                throw "Claude 設定の更新に失敗し、一部の復元にも失敗しました。バックアップを確認してください: $($rollbackErrors -join '; '); original=$originalError"
            }
            throw $originalError
        }

        foreach ($update in $updates) {
            Write-Output "Claude Desktop の設定を更新しました: $($update.Path)"
            if ($update.Backup) { Write-Output "元の設定のバックアップ: $($update.Backup)" }
        }
    } finally {
        foreach ($update in $updates) {
            if (Test-Path -LiteralPath $update.Temp) {
                Remove-Item -LiteralPath $update.Temp -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
if (-not $ConfigureClaudeDesktop) {
    Write-Output 'Claude Desktop の mcpServers.ymm4 に次の内容を設定してください:'
}
Write-Output ($entry | ConvertTo-Json -Depth 3)
Write-Output '接続トークンは YMM4 プラグインの接続ファイルから自動取得します。'
