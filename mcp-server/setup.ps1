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
Write-Output 'Claude Desktop の mcpServers.ymm4 に次の内容を設定してください:'
Write-Output ($entry | ConvertTo-Json -Depth 3)
Write-Output '接続トークンは YMM4 プラグインの接続ファイルから自動取得します。'
