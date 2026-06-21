# YMM4 MCP プラグイン

ClaudeからMCP経由でゆっくりMovieMaker4(YMM4)を操作できるようにするプラグインです。
映像を見ながら・音声を聴きながら、AIが自動でゆっくり実況・解説・茶番・ストーリー動画を生成します。

```
[Claude] ←MCP(stdio)→ [Python MCPサーバー] ←HTTP(8765)→ [YMM4 C#プラグイン] ←内部API→ [YMM4本体]
```

---

## 📁 ファイル構成

```
ymm4プラグイン/
├── YMM4McpPlugin/              # YMM4に読み込まれるC#プラグイン
│   ├── YMM4McpPlugin.csproj
│   ├── McpToolPlugin.cs        # IToolPlugin エントリーポイント
│   ├── McpHttpServer.cs        # HTTPサーバー (port 8765)
│   ├── McpViewModel.cs         # ViewModel (起動/停止UI)
│   ├── McpView.xaml            # コントロールパネルUI
│   ├── McpView.xaml.cs
│   └── BoolToBrushConverter.cs
├── mcp-server/                 # PythonのMCPサーバー
│   ├── server.py               # メインMCPサーバー
│   ├── requirements.txt
│   └── skills/                 # AIスキル集
│       ├── ymm4-jikkyou/       # ゆっくり実況スキル
│       │   └── SKILL.md
│       ├── ymm4-kaisetsu/      # ゆっくり解説スキル
│       │   └── SKILL.md
│       ├── ymm4-chaban/        # ゆっくり茶番スキル
│       │   └── SKILL.md
│       └── ymm4-story/         # ゆっくりストーリースキル
│           └── SKILL.md
├── claude_desktop_config_example.json
└── README.md
```

---

## 🚀 セットアップ手順

### 1. 環境変数の設定

```powershell
[Environment]::SetEnvironmentVariable("YMM4_PATH", "C:\path\to\YukkuriMovieMaker4", "User")
```

### 2. YMM4プラグインのビルド

```powershell
cd YMM4McpPlugin
dotnet build -c Release
# → %YMM4_PATH%\user\plugin\YMM4McpPlugin\YMM4McpPlugin.dll に自動コピー
```

デバッグ時の手動デプロイ：
```powershell
Stop-Process -Name "YukkuriMovieMaker" -Force
Copy-Item "bin\Debug\...\YMM4McpPlugin.dll" "$env:YMM4_PATH\user\plugin\YMM4McpPlugin\" -Force
Start-Process "$env:YMM4_PATH\YukkuriMovieMaker.exe"
```

### 3. PythonのMCPサーバーのセットアップ

```bash
cd mcp-server
pip install -r requirements.txt
```

### 4. Claude Desktopの設定

`%APPDATA%\Claude\claude_desktop_config.json` に追加：

```json
{
  "mcpServers": {
    "ymm4": {
      "command": "python",
      "args": ["C:/Users/ecrea/OneDrive/デスクトップ/ymm4プラグイン/mcp-server/server.py"]
    }
  }
}
```

### 5. YMM4でサーバーを起動

1. YMM4を起動
2. ツールメニュー → 「MCP連携サーバー」を開く
3. **「▶ 起動」** ボタンをクリック
4. `MCPサーバー起動: http://localhost:8765/` が表示されれば完了

---

## 🌐 APIエンドポイント一覧

### プロジェクト・アイテム系
| エンドポイント | 説明 |
|---|---|
| `GET  /api/status` | サーバー生存確認 |
| `GET  /api/project` | プロジェクト情報（FPS・解像度等） |
| `GET  /api/items` | タイムラインの全アイテム取得 |
| `POST /api/project/save` | プロジェクト保存 |
| `POST /api/timeline/duration` | タイムライン長を設定 |

### セリフ・アイテム操作系
| エンドポイント | 説明 |
|---|---|
| `POST /api/voice/add` | VoiceItemを1件追加 |
| `POST /api/script/add` | 複数セリフを一括追加（文字数から長さ自動計算） |
| `POST /api/item/edit` | アイテムのプロパティを変更 |
| `POST /api/item/delete` | アイテムを削除（layer+frame指定） |

### 映像確認系
| エンドポイント | 説明 |
|---|---|
| `GET  /api/preview/capture` | 現在フレームをPNG取得 |
| `POST /api/preview/seek` | 指定フレームへシーク＋PNG取得 |
| `GET  /api/preview/position` | 現在の再生位置取得 |
| `POST /api/preview/export-clip` | ★区間を連番PNG＋音声WAVで書き出す（Gemini解析用） |
| `GET  /api/project/fps` | ★プロジェクトのFPS・解像度を取得（時刻→フレーム変換用） |
| `POST /api/playback/play` | 再生開始 |
| `POST /api/playback/stop` | 再生停止 |

### タイムライン高度操作・状態取得系 ★NEW
| エンドポイント | 説明 |
|---|---|
| `GET  /api/selection` | 現在UIで選択中のアイテム（座標・レイヤー・フレーム・長さ）＋再生位置を取得 |
| `POST /api/items/select` | frame+layer指定でアイテムを選択（clear=trueで全解除） |
| `POST /api/timeline/resolve-overlaps` | レイヤー単位で重なりを解消（gapで最小すき間指定） |
| `POST /api/timeline/shift` | fromFrame以降のアイテムをdeltaフレーム一括シフト |
| `GET  /api/items/effects` | 指定アイテムの全エフェクトとパラメータ現在値を取得 |

> **重なり防止の核心**: `POST /api/items/voice`（およびadd_script）は、`AddVoiceItemAsync`完了後に
> タイムラインを走査して**実際の音声長(`length`/フレーム数)と`endFrame`を取得して返す**ようになりました。
> add_scriptはこの実長を使って次のセリフ開始位置を決めるため、文字数推定のズレによる重なりが根絶されます。

### 全機能アクセス用 汎用API ★NEW
個別エンドポイントで未対応のYMM4内部機能に、リフレクション経由で直接アクセスできます。

| エンドポイント | 説明 |
|---|---|
| `POST /api/command` | 任意のICommandを実行（`name`,`target`,`param`）。UndoCommand等をAPI経由でトリガー |
| `GET  /api/commands` | Main/ActiveTimeline/Player/Projectで利用可能なコマンド一覧と実行可否 |
| `POST /api/reflect/get` | 任意オブジェクトの任意プロパティ/フィールドを取得（`target`,`path`） |
| `POST /api/reflect/set` | 任意プロパティ/フィールドに値を設定（ReactivePropertyの.Valueも対応） |
| `POST /api/reflect/invoke` | 任意メソッドを引数付き呼び出し（戻り値がTaskなら自動await） |
| `GET  /api/reflect/inspect` | オブジェクトの型・プロパティ・メソッド・コマンド一覧（機能の発見用） |

**target**: `Main`(MainViewModel) / `ActiveTimeline` / `Player` / `Project`
**path**: `Items[0].Item.Length` のようにドット・インデックスで深掘り可（ReactivePropertyは自動展開）

### 音声・映像同時取得系
| エンドポイント | 説明 |
|---|---|
| `POST /api/preview/record` | 指定秒数の音声をWAV録音 |
| `POST /api/preview/watch` | 映像キャプチャ＋音声録音を同時実行 ★ |

---

## 🛠️ MCPツール仕様

### `ymm4_interact`（操作系）

| action | sub_action | 説明 |
|---|---|---|
| `get_info` | `status` | サーバー状態確認 |
| `get_info` | `project` | プロジェクト情報取得 |
| `get_info` | `items` | タイムライン全アイテム取得 |
| `get_info` | `effects_list` | エフェクト一覧 |
| `control` | `play` | 再生 |
| `control` | `stop` | 停止 |
| `control` | `save` | 保存 |
| `add_item` | `voice` | セリフ1件追加（実音声長を返す） |
| `add_script` | —— | 複数セリフ一括追加（**実音声長で重なり自動回避**） |
| `edit_item` | `property` | フレーム位置・長さ等を変更 |
| `edit_item` | `delete` | アイテム削除 |
| `edit_item` | `move` | ファイル名指定でアイテムを移動 |
| `edit_item` | `select` | frame+layer指定でアイテムを選択（clearで全解除） |
| `edit_item` | `resolve_overlaps` | 重なり解消（gap指定可） |
| `edit_item` | `shift` | from_frame以降をdeltaフレーム一括シフト |
| `get_info` | `selection` | 選択中アイテムの詳細取得 |
| `get_info` | `commands` | 利用可能コマンド一覧 |
| `get_info` | `effects` | 指定アイテムのエフェクト現在値取得 |
| `control` | `undo` / `redo` | 元に戻す / やり直し |
| `control` | `split` / `align` | 再生位置で分割 / 整列 |

**add_scriptのパラメータ：**
```python
action="add_script",
start_frame=0,       # 開始フレーム
fps=30,              # フレームレート
chars_per_sec=5,     # 話速（実況:5 / 解説:4 / 茶番:6）
lines=[
    {"layer": 7, "character": "ゆっくり霊夢",   "text": "セリフ内容"},
    {"layer": 8, "character": "ゆっくり魔理沙", "text": "セリフ内容"},
]
# 戻り値: total_frames（次のシーンのstart_frame目安）
```

---

### `ymm4_preview`（映像・音声確認系）

| action | 説明 |
|---|---|
| `capture` | 現在フレームをPNG取得 |
| `seek_capture` | 指定フレームへ移動してPNG取得 |
| `position` | 現在の再生位置取得 |
| `record` | 音声のみ録音（WAV保存） |
| `watch` | 映像キャプチャ＋音声録音を同時実行 ★ |

**watchのパラメータ：**
```python
action="watch",
frame=0,                   # 開始フレーム
duration_ms=5000,          # 録音時間（推奨: 最大8000ms）
capture_interval_ms=2000   # 画像取得間隔（推奨: 2000ms以上）
# → 音声: watch_result.wav に保存
# → 画像: PNGとしてレスポンスに含まれる
```

---

### `ymm4_advanced`（全機能アクセス系）★NEW

個別ツールで未対応のYMM4内部機能に、リフレクション経由で直接アクセスする上級ツール。
**YMM4のあらゆる機能をMCPから操作可能**にします。

| action | 説明 |
|---|---|
| `inspect` | 対象オブジェクトのプロパティ・メソッド・コマンド一覧を取得（まず構造を調べる） |
| `get` | 任意プロパティ/フィールドの現在値を取得 |
| `set` | 任意プロパティ/フィールドに値を設定 |
| `invoke` | 任意メソッドを引数付きで呼び出す（Taskは自動await） |
| `command` | 任意のICommandを実行（UIメニュー限定機能を直接トリガー） |
| `list_commands` | 利用可能な全コマンドと実行可否を一覧 |

**使用例：**
```python
# 1. まず構造を調べる
action="inspect", target="ActiveTimeline"

# 2. 値を取得（深掘りパス対応）
action="get", target="ActiveTimeline", path="Items[0].Item.Length"

# 3. 値を設定
action="set", target="Project", path="Name", value="新プロジェクト名"

# 4. メソッド呼び出し
action="invoke", target="ActiveTimeline", method="SelectAll", args=[]

# 5. コマンド実行（元に戻す）
action="command", target="Main", name="UndoCommand"

# 6. どんなコマンドが使えるか発見
action="list_commands"
```

> **設計思想**: `SplitItemCommand`/`AlignItemsCommand` 等の正確なコマンド名はYMM4バージョンで
> 異なる場合があります。`list_commands`/`inspect`で実際の名前を発見してから`command`で実行する
> ワークフローにより、バージョン差異を吸収できます。

---

### `ymm4_analyze_video`（Gemini動画解析系）★NEW

動画を時系列で精密に解析し、「いつ・何が起きたか」（シーン変化／キャラ登場／効果音／テロップ／
動き等）を**タイムスタンプ + YMM4フレーム番号付き**で検出します。
プレビュー画像の単発取得では難しかった「イベント発生フレームの特定」と「音声・映像の同期」を、
Gemini のネイティブ動画理解で実現します。

#### 仕組み（役割分担）
```
[YMM4プレビュー or 元動画]
   ↓ C#: 区間を連番PNG+音声WAVに書き出す / もしくは動画ファイルをそのまま渡す
[Gemini API] 映像+音声を時系列解析 → タイムスタンプ付きイベントJSON
   ↓ Python: time_sec を FPS でフレーム番号に変換
[Claude] イベントに合わせてセリフ生成・add_scriptで配置
```
全フレームをClaudeに投げるのではなく、**重い解析はGeminiに集約**し、Claudeは構造化済みの
イベントリストを受け取って意味づけ・台本化に専念します。

#### 2つのモード
| source | 解析対象 | 主なパラメータ |
|---|---|---|
| `preview` | YMM4プレビューの指定フレーム区間（配置済み素材） | `start_frame` / `end_frame` / `step_frames` / `record_audio` |
| `file` | 取り込み前の元動画ファイル(mp4等) | `video_path` / `base_frame` / `fps` |

#### 検出イベントのカスタマイズ
`event_instruction` で検出したいイベントを自由に指定できます（**未指定なら全イベント検出**）。
```python
# 全部検出（デフォルト）
action source="preview", start_frame=0, end_frame=600

# 効果音とキャラ登場だけに絞る
event_instruction="効果音が鳴った瞬間と、新しいキャラが画面に出た瞬間だけ検出して"

# 元動画ファイルを解析し、フレーム300以降に対応づける
source="file", video_path="C:/movie/boss.mp4", base_frame=300
```

#### 戻り値（例）
```json
{
  "success": true,
  "summary": "ボス戦の動画。中盤で敵が出現し、終盤に撃破される。",
  "events": [
    {
      "time_sec": 12.3, "end_sec": 14.0,
      "frame": 369, "end_frame": 420,          // ← FPSから自動変換されたYMM4フレーム
      "type": "character", "label": "ボス出現",
      "description": "画面中央に大型の敵が出現", "audio": "重低音のSE",
      "importance": 5
    }
  ],
  "fps_used": 30, "base_frame": 0
}
```
`frame` がそのまま `add_script` の `start_frame` 等に使えるため、**映像イベントに同期したセリフ配置**が可能です。

#### セットアップ
1. `pip install google-genai`（`requirements.txt` に追加済み）
2. [Google AI Studio](https://aistudio.google.com/) でAPIキーを発行
3. `claude_desktop_config.json` の `env` に `GEMINI_API_KEY` を設定（`claude_desktop_config_example.json` 参照）

> モデルは既定 `gemini-2.5-flash`。高精度が必要なら `model="gemini-2.5-pro"` を指定。
> APIキー未設定/SDK未導入でもサーバーは起動し、解析ツール呼び出し時に案内メッセージを返します。

---

## 🤖 AIスキル一覧

`mcp-server/skills/` 以下に動画ジャンル別のスキルがあります。
Claudeはユーザーの依頼内容に応じて自動的に対応するスキルを参照します。

### キャラクター役割

| キャラ | 実況 | 解説 | 茶番 | ストーリー |
|---|---|---|---|---|
| **霊夢** (L7) | ボケ・マイペース | 生徒・質問役 | ボケ・天然 | 主人公・感情表現 |
| **魔理沙** (L8) | ツッコミ・解説 | 解説役・先生 | ツッコミ・進行 | 相棒・行動派 |

### スキル詳細

| スキル | ファイル | 話速 | 特徴 |
|---|---|---|---|
| **ゆっくり実況** | `ymm4-jikkyou/SKILL.md` | 5文字/秒 | watchで映像確認しながらセリフ生成。ゲームイベントに合わせたテンプレあり |
| **ゆっくり解説** | `ymm4-kaisetsu/SKILL.md` | 4文字/秒 | 魔理沙が解説・霊夢が質問。ポイント3〜5個の構成テンプレあり |
| **ゆっくり茶番** | `ymm4-chaban/SKILL.md` | 6文字/秒 | 霊夢（ボケ）と魔理沙（ツッコミ）。ツッコミは3〜5f後に即返す |
| **ゆっくりストーリー** | `ymm4-story/SKILL.md` | 4〜6文字/秒 | ナレーター(L6)追加。感動/ホラー/ファンタジー/ミステリー対応 |

---

## 💬 Claudeへの指示例

### 実況動画
```
マインクラフトのボス戦動画を見ながらゆっくり実況を作って
```

### 解説動画
```
スケルトンキングの攻略法をゆっくり解説動画にして
魔理沙に解説させて、霊夢に質問させて
```

### 茶番動画
```
霊夢と魔理沙で買い物に行く茶番コントを作って
```

### ストーリー動画
```
霊夢と魔理沙が謎の洞窟を探索する短編ミステリーを作って
```

### タイムライン確認・編集
```
今のタイムラインを確認して、セリフが被っているところを修正して
100フレームごとに映像と音声を確認して
```

---

## 📋 標準制作ワークフロー

```
① watchで映像を見る＋音を聴いてシーンを把握
        ↓
② シーン表と台本を設計
        ↓
③ add_scriptでシーンごとに一括配置
        ↓
④ 100fごとにseek_captureで映像・セリフを確認
        ↓
⑤ ズレ・被りをedit_itemで修正
        ↓
⑥ watchで最終確認（音声RMS・映像同期）
        ↓
⑦ 完成！プロジェクト保存
```

---

## ⚠️ 注意事項・既知の制限

| 項目 | 内容 |
|---|---|
| YMM4バージョン | v4.35以降（.NET 9 / .NET 10対応）が必要 |
| watchの上限 | 画像+音声データが大きいため最大8秒程度推奨 |
| 初回キャプチャ | シーク直後の1枚目は黒になることがある（描画待ち）。正常動作 |
| ポート競合 | 8765が競合する場合は `McpHttpServer.cs` の `Port` を変更 |
| 内部API | `IMainViewModel` の実装はYMM4バージョンにより異なる場合あり |
