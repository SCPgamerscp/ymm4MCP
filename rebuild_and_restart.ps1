# YMM4プラグイン 完全自動 再ビルド＆再起動スクリプト
# 使い方:
#   .\rebuild_and_restart.ps1           → YMM4終了→ビルド→起動
#   .\rebuild_and_restart.ps1 -NoCopy   → YMM4起動したままビルドのみ（DLLはbinフォルダへ）

param(
    [switch]$NoCopy  # YMM4を終了せずビルドのみ
)

if ([string]::IsNullOrWhiteSpace($env:YMM4_PATH)) {
    throw "環境変数 YMM4_PATH にYukkuriMovieMaker4のインストール先を設定してください。"
}

$repoRoot = $PSScriptRoot
$ymm4Root = [IO.Path]::GetFullPath($env:YMM4_PATH)
$ymm4exe = Join-Path $ymm4Root "YukkuriMovieMaker.exe"
$proj = Join-Path $repoRoot "YMM4McpPlugin\YMM4McpPlugin.csproj"
$pluginDst = Join-Path $ymm4Root "user\plugin\YMM4McpPlugin"
$dllSrc = Join-Path $repoRoot "YMM4McpPlugin\bin\Debug\net10.0-windows10.0.19041.0\YMM4McpPlugin.dll"
$connectionFile = Join-Path $env:LOCALAPPDATA "YMM4MCP\connection.json"

foreach ($requiredPath in @($ymm4exe, $proj)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "必要なファイルが見つかりません: $requiredPath"
    }
}

if ($NoCopy) {
    # YMM4起動中はコピー不要でビルドのみ
    Write-Host "[ビルドのみモード]" -ForegroundColor Yellow
    Write-Host "[1/1] ビルド中..." -ForegroundColor Cyan
    $result = & dotnet build $proj "-p:YMM4LibPath=$ymm4Root" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ビルド失敗！" -ForegroundColor Red
        $result | Where-Object { $_ -match "error" } | Select-Object -Last 10
    } else {
        Write-Host "ビルド成功！DLL: $dllSrc" -ForegroundColor Green
        Write-Host "→ YMM4を再起動してから手動でプラグインを更新してください" -ForegroundColor Yellow
    }
    Read-Host "Enterで終了"
    exit
}

# === 通常モード: YMM4終了→ビルド→コピー→起動 ===

# 1. YMM4終了
Write-Host "[1/3] YMM4を終了中..." -ForegroundColor Cyan
$p = Get-Process -Name "YukkuriMovieMaker" -ErrorAction SilentlyContinue
if ($p) { $p | Stop-Process -Force; Start-Sleep 2; Write-Host "  -> 終了完了" }
else    { Write-Host "  -> 起動していませんでした" }

# 2. ビルド
Write-Host "[2/3] ビルド中..." -ForegroundColor Cyan
$result = & dotnet build $proj "-p:YMM4LibPath=$ymm4Root" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  -> ビルド失敗！" -ForegroundColor Red
    $result | Where-Object { $_ -match "error CS" } | Select-Object -Last 10
    Read-Host "Enterで終了"
    exit 1
}
Write-Host "  -> ビルド成功" -ForegroundColor Green

# DLLを手動コピー（post-buildがファイルロックで失敗する場合に備え）
New-Item -ItemType Directory -Path $pluginDst -Force | Out-Null
Copy-Item $dllSrc $pluginDst -Force
Write-Host "  -> DLLコピー完了: $pluginDst" -ForegroundColor Green

# 3. YMM4起動
Write-Host "[3/3] YMM4を起動中..." -ForegroundColor Cyan
Start-Process $ymm4exe
Write-Host "  -> 起動完了。MCPサーバーの起動を待機中..."

for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep 2
    try {
        if (-not (Test-Path -LiteralPath $connectionFile)) {
            throw "接続情報を待機中"
        }
        $connection = Get-Content -LiteralPath $connectionFile -Raw | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace($connection.api_base) -or [string]::IsNullOrWhiteSpace($connection.token)) {
            throw "接続情報が不完全です"
        }
        $headers = @{ "X-Ymm4-Token" = $connection.token }
        $statusUri = $connection.api_base.TrimEnd('/') + "/status"
        $r = Invoke-RestMethod -Uri $statusUri -Headers $headers -TimeoutSec 2
        Write-Host ""
        Write-Host "MCPサーバー起動確認！ ($($r.timestamp))" -ForegroundColor Green
        Write-Host "URL: $($connection.api_base)/" -ForegroundColor Green
        Write-Host ""
        Write-Host "完了！" -ForegroundColor Green
        exit 0
    } catch {
        Write-Host "  待機中... $(($i+1)*2)秒"
    }
}

Write-Host "タイムアウト。YMM4のツール→MCP連携サーバーを確認してください。" -ForegroundColor Yellow
Read-Host "Enterで終了"
