param([switch]$ConfigureClaudeDesktop, [switch]$CheckConnection)
$ErrorActionPreference = 'Stop'
$serverDir = $PSScriptRoot
$venvDir = Join-Path $serverDir '.venv'
$venvPython = Join-Path $venvDir 'Scripts\python.exe'

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 -m venv $venvDir
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    & python -m venv $venvDir
} else {
    throw 'Python 3 が見つかりません。Python をインストールしてから再実行してください。'
}
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython)) {
    throw 'Python の仮想環境を作成できませんでした。'
}

& $venvPython -m pip install --disable-pip-version-check -r (Join-Path $serverDir 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'MCP サーバーの依存関係をインストールできませんでした。' }

$entry = @{ command = $venvPython; args = @((Join-Path $serverDir 'server.py')) }
if ($ConfigureClaudeDesktop) {
    if (-not $env:APPDATA) { throw 'APPDATA が見つかりません。Claude Desktop の設定場所を確認してください。' }
    $configPath = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
    $configDir = Split-Path -Parent $configPath
    [IO.Directory]::CreateDirectory($configDir) | Out-Null
    if (Test-Path -LiteralPath $configPath) {
        $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $config -or $config -isnot [pscustomobject]) {
            throw 'Claude Desktop の設定は JSON オブジェクトである必要があります。'
        }
        $backupPath = "$configPath.bak.$([guid]::NewGuid().ToString('N'))"
        Copy-Item -LiteralPath $configPath -Destination $backupPath
    } else {
        $config = [pscustomobject]@{}
        $backupPath = $null
    }
    if ($null -eq $config.mcpServers) {
        $config | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{}) -Force
    } elseif ($config.mcpServers -isnot [pscustomobject]) {
        throw 'mcpServers は JSON オブジェクトである必要があります。'
    }
    $config.mcpServers | Add-Member -NotePropertyName ymm4 -NotePropertyValue $entry -Force
    $tempPath = "$configPath.tmp.$([guid]::NewGuid().ToString('N'))"
    try {
        $json = $config | ConvertTo-Json -Depth 100
        [IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $configPath) {
            [IO.File]::Replace($tempPath, $configPath, $null)
        } else {
            [IO.File]::Move($tempPath, $configPath)
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath }
    }
    Write-Output "Claude Desktop の設定を更新しました: $configPath"
    if ($backupPath) { Write-Output "元の設定のバックアップ: $backupPath" }
}
if (-not $ConfigureClaudeDesktop) {
    Write-Output 'Claude Desktop の mcpServers.ymm4 に次の内容を設定してください:'
}
Write-Output ($entry | ConvertTo-Json -Depth 3)
Write-Output '接続トークンは YMM4 プラグインの接続ファイルから自動取得します。'
if ($CheckConnection) {
    & $venvPython (Join-Path $serverDir 'check_connection.py')
    if ($LASTEXITCODE -ne 0) { throw 'YMM4プラグインへの接続確認に失敗しました。' }
}
