
# 追加後に実際のアイテム長を取得して次フレームを決めるスクリプト

$lines = @(
    @{character="ゆっくり霊夢"; text="みんな～！今日はMCPについて解説するわよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="MCPって何の略だぜ？"; layer=2},
    @{character="ゆっくり霊夢"; text="Model Context Protocolよ。AIとツールを繋ぐ仕組みのことよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="ふむふむ。もっと詳しく教えてくれ！"; layer=2},
    @{character="ゆっくり霊夢"; text="簡単に言うと、AIがファイルを読んだり、アプリを操作したりできるようになる規格よ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="それはすごいな！まるでAIに手足が生えるみたいだぜ！"; layer=2},
    @{character="ゆっくり霊夢"; text="その通り！次はMCPのアーキテクチャを見ていくわよ。"; layer=1},
    @{character="ゆっくり魔理沙"; text="アーキテクチャ？難しそうだな…"; layer=2},
    @{character="ゆっくり霊夢"; text="大丈夫よ！登場人物は3つだけ。MCPホスト、MCPクライアント、MCPサーバーよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="ホストってClaudeとかのこと？"; layer=2},
    @{character="ゆっくり霊夢"; text="そうよ！AIアシスタントがホストで、その中にクライアントが入っているの。"; layer=1},
    @{character="ゆっくり魔理沙"; text="じゃあMCPサーバーは？"; layer=2},
    @{character="ゆっくり霊夢"; text="ツールやデータを提供する側ね。ファイルシステムやAPIを繋ぐの！"; layer=1},
    @{character="ゆっくり魔理沙"; text="なるほど！じゃあ3つのプリミティブってのは？"; layer=2},
    @{character="ゆっくり霊夢"; text="MCPには3つの基本要素があるの。ツール、リソース、プロンプトよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="ツールはさっき言ってた手足の部分だな！"; layer=2},
    @{character="ゆっくり霊夢"; text="正解！AIが呼び出せる関数みたいなものよ。検索したりファイルを操作したり！"; layer=1},
    @{character="ゆっくり魔理沙"; text="リソースとプロンプトはどう違うんだぜ？"; layer=2},
    @{character="ゆっくり霊夢"; text="リソースはファイルやデータベースなどの読み取り専用データ。プロンプトはテンプレートよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="わかってきたぜ！最後は開発スタックだな！"; layer=2},
    @{character="ゆっくり霊夢"; text="MCPサーバーはPythonかTypeScriptで簡単に作れるわよ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="このYMM4プラグインもMCPサーバーなのか！"; layer=2},
    @{character="ゆっくり霊夢"; text="そうよ！C#で書いたHTTPサーバーをPythonのMCPブリッジが繋いでいるの！"; layer=1},
    @{character="ゆっくり魔理沙"; text="ClaudeがYMM4をMCP経由で操作してるってことか！すごいな！"; layer=2},
    @{character="ゆっくり霊夢"; text="MCPを使えばAIがあなたの作業を強力にサポートしてくれるわ！"; layer=1},
    @{character="ゆっくり魔理沙"; text="みんなもMCPを活用して、動画制作を自動化してみてくれ！以上、ゆっくりしていってね！"; layer=2}
)

$frame = 0
$success = 0

function SendVoice($character, $text, $frame, $layer) {
    $body = [System.Text.Encoding]::UTF8.GetBytes(
        (ConvertTo-Json @{character=$character; text=$text; frame=$frame; layer=$layer} -Compress)
    )
    $req = [System.Net.WebRequest]::Create("http://localhost:8765/api/items/voice")
    $req.Method = "POST"
    $req.ContentType = "application/json; charset=utf-8"
    $req.ContentLength = $body.Length
    $req.Timeout = 30000
    $s = $req.GetRequestStream(); $s.Write($body,0,$body.Length); $s.Close()
    $res = $req.GetResponse()
    return (New-Object System.IO.StreamReader($res.GetResponseStream(),[System.Text.Encoding]::UTF8)).ReadToEnd() | ConvertFrom-Json
}

function GetActualLength($frame, $layer) {
    # 追加したアイテムの実際のlengthをAPIから取得
    $r = Invoke-RestMethod -Uri "http://localhost:8765/api/items" -TimeoutSec 5
    foreach ($item in $r.items) {
        if ($item.frame -eq $frame -and $item.layer -eq $layer) {
            return $item.length
        }
    }
    return $null
}

foreach ($line in $lines) {
    Write-Host "追加中 [frame=$frame] $($line.character): $($line.text.Substring(0,[Math]::Min(25,$line.text.Length)))..."
    
    $res = SendVoice $line.character $line.text $frame $line.layer
    
    if ($res.success) {
        # 実際のアイテム長を取得
        Start-Sleep -Milliseconds 300  # 音声生成完了を少し待つ
        $actualLen = GetActualLength $frame $line.layer
        
        if ($actualLen -and $actualLen -gt 0) {
            Write-Host "  -> OK! 実際のlength=$actualLen フレーム"
            $frame += $actualLen
        } else {
            # 取得できなかった場合は文字数から推定（フォールバック）
            $estimated = [Math]::Max(90, [int]($line.text.Length / 5.0 * 30))
            Write-Host "  -> OK! length取得失敗、推定=$estimated フレーム"
            $frame += $estimated
        }
        $success++
    } else {
        Write-Host "  -> 失敗: $($res.error)"
        $frame += 90
    }
}

Write-Host ""
Write-Host "完了！ 成功=$success / $($lines.Count)  総フレーム=$frame"
