# YMM4 MCP プラグイン

ClaudeからMCP経由でゆっくりMovieMaker4(YMM4)を操作できるようにするプラグインです。
映像を見ながら・音声を聴きながら、AIが自動でゆっくり実況・解説・茶番・ストーリー動画を生成します。

```
[Claude] ←MCP(stdio)→ [Python MCPサーバー] ←HTTP(8765)→ [YMM4 C#プラグイン] ←内部API→ [YMM4本体]
```

---

## 🚀 はじめに（Windows）

1. [Releases](https://github.com/SCPgamerscp/ymm4MCP/releases) から最新版の `YMM4McpPlugin-*.ymme` をダウンロードします。YMM4を閉じて `.ymme` をダブルクリックし、YMM4の案内に従ってインストールします。利用者は `YMM4_PATH` や .NET SDK を設定する必要はありません。
2. YMM4を起動します。新規インストールではプラグインのHTTPサーバーが自動起動します。「ツール」→「MCP連携サーバー」で状態を確認できます。既に自動起動を無効にしている場合は、この画面の「▶ 起動」を押します。
3. PythonのMCPサーバー用に、このリポジトリの[ソースコード](https://github.com/SCPgamerscp/ymm4MCP)を取得し、`mcp-server` で依存関係をインストールします（Pythonが必要です）。`.ymme` にPythonサーバーは含まれません。

   ```powershell
   cd C:\path\to\ymm4MCP\mcp-server
   python -m pip install -r requirements.txt
   ```

4. Claude Desktopの `%APPDATA%\Claude\claude_desktop_config.json` にMCPサーバーを追加します。`args` は取得したソースコードの実際の場所に置き換えてください。

   ```json
   {
     "mcpServers": {
       "ymm4": {
         "command": "python",
         "args": ["C:/path/to/ymm4MCP/mcp-server/server.py"]
       }
     }
   }
   ```

プラグインは `%LOCALAPPDATA%\YMM4MCP\connection.json` に接続先と秘密トークンを書き込み、Python側は毎回これを読み取ります。このファイルを共有しないでください。接続確認には、YMM4を起動した状態で次を実行します（`/api/status` にもトークンが必要です）。

```powershell
$connection = Get-Content "$env:LOCALAPPDATA\YMM4MCP\connection.json" -Raw | ConvertFrom-Json
Invoke-RestMethod -Uri "$($connection.api_base)/status" -Headers @{ 'X-Ymm4-Token' = $connection.token }
```

ポートはツール画面でサーバーを停止してから変更・保存し、再度起動します。Python側は更新された接続情報から新しいポートを自動取得します。手動で接続先を指定する場合だけ `YMM4_API_BASE=http://127.0.0.1:<port>/api` を設定してください。

---

## 📁 ファイル構成

```
ymm4プラグイン/
├── YMM4McpPlugin/              # YMM4に読み込まれるC#プラグイン
│   ├── YMM4McpPlugin.csproj
│   ├── McpToolPlugin.cs        # IToolPlugin エントリーポイント
│   ├── McpHttpServer.cs        # HTTPサーバー (port 8765)
│   ├── McpJobs.cs              # ジョブ状態・cancel/resume
│   ├── McpExport.cs            # 完成動画書き出し・プロジェクトopen/save-as
│   ├── McpEdits.cs             # EditPlan状態・idempotencyバインディング・checkpoint
│   ├── McpKeyframes.cs         # Animationキーフレーム
│   ├── McpEditing.cs           # 安定ID・revision・素材追加
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

## 開発者向け：ビルドと配布

`YMM4_PATH` はソースからビルドするときにだけ必要です。YMM4のインストール先を指定してください。

```powershell
$env:YMM4_PATH = "C:\path\to\YukkuriMovieMaker4"
dotnet build YMM4McpPlugin/YMM4McpPlugin.csproj -c Release -p:CreateYmme=true
# → artifacts\YMM4McpPlugin-1.5.0.ymme
```

PRではWindowsのGitHub Actionsが公式YMM4 LiteのDLLを参照して `.ymme` を検証します。`v1.5.0` タグを `main` のコミットに付けると、バージョン一致を確認してGitHub Releaseへ添付します。YMM4本体のDLLは `.ymme` に含めません。

---

## 🌐 APIエンドポイント一覧

### プロジェクト・アイテム系
| エンドポイント | 説明 |
|---|---|
| `GET  /api/status` | サーバー生存確認 |
| `GET  /api/project` | プロジェクト情報（FPS・解像度等） |
| `GET  /api/items` | タイムラインの全アイテム取得 |
| `POST /api/project/save` | プロジェクト保存 |
| `POST /api/project/open` | `.ymmp` をパス指定で開く |
| `POST /api/project/save-as` | 別名保存（`overwrite`で上書き）。既存ファイルは先にバックアップ |
| `POST /api/project/export` | 完成動画書き出しを**ジョブとして投入**。すぐ `job_id` を返す |
| `GET  /api/jobs` | 最近のジョブ一覧 |
| `GET  /api/jobs/{id}` | 進捗・phase・検証結果 |
| `POST /api/jobs/{id}/cancel` | キャンセル要求 |
| `POST /api/jobs/{id}/resume` | 失敗/中断ジョブを再投入 |
| `GET  /api/edits/state` | 現在のタイムラインをEditPlan構造（1シーン）で取得。適用済みバインディング含む |
| `GET  /api/edits/bindings` | 保存済み `idempotency_key` → item_id 対応 |
| `POST /api/edits/bindings` | apply後の plan_id と item_id の対応を保存 |
| `POST /api/edits/checkpoint` | 現在の `item_id` スナップショット。`backup=true` なら保存済み `.ymmp` をコピー |
| `GET  /api/edits/checkpoints` | チェックポイント一覧 |
| `POST /api/edits/rollback` | スナップショット以降に追加されたアイテムを削除（削除済みの復元はしない） |
| `POST /api/timeline/duration` | タイムライン長を設定 |

### セリフ・アイテム操作系
| エンドポイント | 説明 |
|---|---|
| `POST /api/items/voice` | VoiceItemを1件追加し、実際の長さ・`item_id`・`revision`を返す |
| MCP `add_script` | 複数セリフを一括追加（HTTPの独立エンドポイントではありません） |
| `POST /api/items/prop` | `item_id`（推奨）またはframe+layerでプロパティを変更。`expected_revision`対応 |
| `GET  /api/items/keyframes` | Animationプロパティのキーフレーム一覧。`item_id`またはframe+layer、任意で`prop` |
| `POST /api/items/keyframe` | キーフレームの set / remove / clear。`item_id`+`expected_revision`対応 |
| `POST /api/items/delete` | `item_id` / `item_ids`（推奨）またはframe+layer/layersで削除。`expected_revision`対応 |

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
| `POST /api/items/select` | `item_id`（推奨）またはframe+layer指定でアイテムを選択（clear=trueで全解除） |
| `POST /api/timeline/resolve-overlaps` | レイヤー単位で重なりを解消（gapで最小すき間指定） |
| `POST /api/timeline/shift` | fromFrame以降のアイテムをdeltaフレーム一括シフト |
| `GET  /api/items/effects` | 指定アイテムの全エフェクトとパラメータ現在値を取得 |
| `GET  /api/commands` | 利用可能なコマンド一覧と実行可否（読み取り専用。高度APIを有効にしなくても利用可） |

> **重なり防止の核心**: `POST /api/items/voice`（およびadd_script）は、`AddVoiceItemAsync`完了後に
> タイムラインを走査して**実際の音声長(`length`/フレーム数)と`endFrame`を取得して返す**ようになりました。
> add_scriptはこの実長を使って次のセリフ開始位置を決めるため、文字数推定のズレによる重なりが根絶されます。

#### 安定したアイテム参照と競合防止

`GET /api/items` と各追加APIは `item_id`、`revision`、`identity_persistent` を返します。
`/api/items/prop` と `/api/items/delete` と `/api/items/keyframe` では `item_id` を指定するとframe/layerより優先され、
さらに同じリクエストで `expected_revision` を渡すと、取得後に別操作で変更されたアイテムを誤って上書きしません。
`expected_revision` は対象を曖昧にしないため `item_id` と必ず組み合わせてください。
不一致時は `REVISION_CONFLICT` となるため、最新のitemsを取得してから操作を組み直してください。
フレーム・レイヤー・長さ・すき間などの整数パラメータはMCPサーバーで事前検証され、
負のフレーム/レイヤー、0以下の長さ、範囲外の値はYMM4へ送信せずエラーになります。
一括削除で従来から使われる `frame=-1` / `layer=-1` の省略指定には引き続き対応します。
YMM4のHTTP APIへ直接送る場合も、指定した整数の型・範囲と `layers` の各要素を検証します。
不正な入力は編集前にHTTP 400の `INVALID_ARGUMENT` を返します。省略した整数だけ既定値を使います。
`move` は空でない `filename` と有効な配置が必要です。`shift` と `resolve-overlaps` は対象全件の
移動先を先に検証し、1件でもフレーム範囲外なら全件を変更せず拒否します（負方向への範囲外シフトも同様）。
空のレイヤーを含むレイヤー自体の実在判定は、現在のアイテム一覧だけでは行いません。

```json
{
  "item_id": "native:...",
  "expected_revision": "現在のrevision",
  "prop": "Length",
  "value": "120"
}
```

YMM4本体がネイティブ識別子を公開するアイテムは再起動後も同じIDになります。
ネイティブ識別子がない型は実行中のみ安定する `runtime:` IDとなり、`identity_persistent` が `false` になります。
そのIDは再起動をまたいで保存せず、再接続後に取り直してください。

#### ジョブと完成動画の書き出し

長い処理は1リクエストで待たず、ジョブIDを返して進捗を取りにいきます。MCPを切断してもYMM4側のジョブは続きます。YMM4プロセス自体が終了すると `interrupted` になり、`resume` で再投入できます（途中フレームからの再開ではありません）。

```json
POST /api/project/export
{
  "path": "C:/Videos/final.mp4",
  "format": "mp4",
  "overwrite": false,
  "timeout_seconds": 1800,
  "idempotency_key": "episode-12-render"
}
```

戻り値の `job_id` を `GET /api/jobs/{id}` でポーリングします。完了時はファイルの存在・サイズを確認し、MP4 は `ftyp`、`moov`、`mdat`、映像トラック、長さ、解像度を検証して `duration_seconds`、`width`、`height`、`has_audio` を返します。WAV は `RIFF/WAVE` ヘッダを確認します。コンテナの構造検証であり、映像・音声を実際にデコードして再生品質まで判定するものではありません。

同じ `idempotency_key` の再送は、実行中または成功済みのジョブをそのまま返します。既存ファイルを消したくない場合は `overwrite` を省略してください。

本体の出力APIはバージョン差があるため実行時に探索します。パス付きメソッドが無い場合は `EXPORT_METHOD_UNAVAILABLE` または `EXPORT_DIALOG_REQUIRED` になり、ダイアログ操作は自動化しません。

MCPからは次のように呼びます。

```python
action="control", sub_action="export", path="C:/Videos/final.mp4"
action="get_info", sub_action="job", job_id="job_..."
action="control", sub_action="cancel_job", job_id="job_..."
action="control", sub_action="open", path="C:/proj/a.ymmp"
action="control", sub_action="save_as", path="C:/proj/b.ymmp", overwrite=True
```

`open` は未保存変更がある場合 `UNSAVED_CHANGES`、保存状態を取得できない場合 `PROJECT_SAVE_STATE_UNKNOWN` を返して停止します。現在の内容を確認して保存した後に再実行するか、破棄を意図する場合は `force=True` を指定してください。YMM4本体の確認ダイアログが表示される場合は、その操作も必要です。

`save` と `save_as` で既存の `.ymmp` を上書きする前に、以前の内容をローカルの `YMM4MCP/project-backups` にコピーします。応答の `backup_path` から退避先を確認できます。コピーに失敗した場合は `BACKUP_FAILED` で保存せず停止します。新規保存または現在のプロジェクトパスを取得できない場合はバックアップ先が `null` です。退避ファイルの整理は利用者が行ってください。

`get_info/audio_qa` は YMM4 側の絶対パス `path` にある PCM 16-bit WAV（モノラル/ステレオ）を検査します。`min_silence_seconds`（既定2秒）以上の無音、フルスケール付近のサンプルの継続、左右チャンネルの大きな RMS 差を `issues` に返します。`passed` は error が無いときだけ true です。対応外の形式は `AUDIO_FORMAT_UNSUPPORTED` を返します。映像に埋め込まれた音声や MP3 はこの検査の対象外です。

#### 宣言的EditPlan（dry-run / 差分適用 / シーン単位transaction）

LLMが数百回の低レベルAPIを直接組み立てる代わりに、完成状態を渡して差分だけ適用します。同じ `idempotency_key` の再送は、同じ計画の対象アイテムが変更されずに残っていれば二重追加しません。削除されたアイテムは `reconcile_edit` で補えます。既存の計画外アイテムは削除しません。

`apply_edit` はシーンをtransactionにします。あるシーンの追加が失敗すると、**そのシーンで追加したアイテムだけ削除**し、先に完了したシーンは残します。`atomic_scenes=false` で旧来どおり部分追加を残すこともできます。通信失敗（追加できたか不明）では自動削除しません。

同じキーで計画を変えると `IDEMPOTENCY_KEY_CONFLICT`、適用済みアイテムの revision が変わると `EDIT_STATE_CONFLICT` で編集前に停止します。全シーンが通ったあとは `/api/items` の実状態を照合し、不一致や取得失敗は `EDIT_VERIFY_FAILED`、バインディング保存失敗は `BINDINGS_NOT_SAVED` を返します。これらは追加済みアイテムを巻き戻さないため、`details` と最新の `items` を確認してから再実行してください。再適用情報を読み取れない場合は `BINDINGS_UNAVAILABLE` で編集せず停止します。

チェックポイントは `item_id` のスナップショットです。rollbackは「後から増えたアイテムの削除」であり、削除済みアイテムやプロパティ変更は復元しません。完全復元が必要なら `backup=true` で `.ymmp` をコピーし、`control/open` で開きます。

```python
plan = {
  "project": {"fps": 30},
  "scenes": [
    {"id": "intro", "items": [
      {"id": "bg-1", "type": "video", "source": "C:/Videos/play.mp4", "layer": 0, "frame": 0},
      {"id": "line-1", "type": "dialogue", "character": "ゆっくり霊夢", "text": "こんにちは", "layer": 7},
      {"id": "cap-1", "type": "subtitle", "text": "こんにちは", "layer": 9, "length": 90},
    ]}
  ]
}
action="plan_edit", plan=plan, idempotency_key="episode-1"   # 変更なし。差分と警告だけ
action="control", sub_action="checkpoint", reason="before apply", backup=True
action="apply_edit", plan=plan, idempotency_key="episode-1"  # シーン失敗時はそのシーンだけrollback
action="reconcile_edit", plan=plan, idempotency_key="episode-1"
action="control", sub_action="rollback", checkpoint_id="cp_..."
action="get_info", sub_action="edit_state"
```

対応タイプ: `video` / `audio` / `bgm` / `se` / `image` / `text` / `subtitle` / `dialogue` / `voice` / `tachie` / `face`。
映像・音声QAと修正回数制限はこのスライスの対象外です。

#### キーフレーム

立ち絵や字幕の X / Y / Zoom / Opacity / Volume などを、アイテム開始からの相対フレーム `at` で時間変化させます。

```json
{
  "item_id": "native:...",
  "expected_revision": "現在のrevision",
  "prop": "X",
  "action": "set",
  "at": 0,
  "value": -400
}
```

MCPからは `edit_item` / `keyframe`（`keyframe_action=set|remove|clear`）と `get_info` / `keyframes` で呼びます。
YMM4内部の Animation API をリフレクションで叩くため、対象バージョンでメソッドが無い場合は `KEYFRAME_METHOD_UNAVAILABLE` になります。そのときは `inspect` で署名を確認してください。

### 全機能アクセス用 汎用API ★NEW
個別エンドポイントで未対応のYMM4内部機能に、リフレクション経由で直接アクセスできます。

| エンドポイント | 説明 |
|---|---|
| `POST /api/command` | 任意のICommandを実行（`name`,`target`,`param`）。UndoCommand等をAPI経由でトリガー |
| `POST /api/reflect/get` | 任意オブジェクトの任意プロパティ/フィールドを取得（`target`,`path`） |
| `POST /api/reflect/set` | 任意プロパティ/フィールドに値を設定（ReactivePropertyの.Valueも対応） |
| `POST /api/reflect/invoke` | 任意メソッドを引数付き呼び出し（戻り値がTaskなら自動await） |
| `GET  /api/reflect/inspect` | オブジェクトの型・プロパティ・メソッド・コマンド一覧（機能の発見用） |

`GET /api/effects/describe?name=...` は `effects/list` のエフェクト名または完全型名を受け取り、公開プロパティの型・書き込み可否と、属性に記述された範囲・既定値・表示名・単位を返します。属性がない値は `null` です。MCPでは `get_info` の `sub_action="effect_metadata"` と `name` で取得できます。YMM4の実行状態に依存する適用可能対象はこの情報だけでは判定しません。

**target**: `Main`(MainViewModel) / `ActiveTimeline` / `Player` / `Project`
**path**: `Items[0].Item.Length` のようにドット・インデックスで深掘り可（ReactivePropertyは自動展開）
`reflect/get` は存在しないメンバー・範囲外または不正な添字に `PATH_NOT_FOUND` を返します。存在するメンバーの値が `null` なら成功応答で `value: null` を返します。

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
| `get_info` | `project` | プロジェクト情報と未保存変更の状態を取得 |
| `get_info` | `items` | タイムライン全アイテム取得 |
| `get_info` | `media` | YMM4側の素材ファイルの存在・サイズ・拡張子・最終更新日時を取得。絶対パス `path` を指定 |
| `get_info` | `effects_list` | エフェクト一覧 |
| `get_info` | `effect_metadata` | `name` でエフェクトの公開設定項目と属性メタデータを取得 |
| `control` | `play` | 再生 |
| `control` | `stop` | 停止 |
| `control` | `save` | 保存 |
| `control` | `open` / `save_as` | プロジェクトをパス指定で開く / 別名保存 |
| `control` | `export` | 完成動画書き出しをジョブ投入。`path` 必須 |
| `get_info` | `jobs` / `job` | ジョブ一覧 / `job_id` の進捗 |
| `get_info` | `edit_state` | 現在のタイムラインをEditPlan構造で取得 |
| `get_info` | `checkpoints` | チェックポイント一覧。`checkpoint_id` で1件 |
| `control` | `cancel_job` / `resume_job` | ジョブ中止 / 失敗・中断の再投入 |
| `control` | `checkpoint` / `rollback` | item_idスナップショット / 追加分の削除。`backup`で.ymmpコピー |
| `plan_edit` | —— | 完成状態のEditPlanを検証し、差分と警告だけ返す（編集なし） |
| `apply_edit` | —— | 差分だけ適用。シーン失敗時はそのシーンの追加分をrollback。同じ `idempotency_key` は二重追加しない |
| `reconcile_edit` | —— | 中断後に不足分だけ再実行 |
| `qa_gate` | —— | 現在のタイムラインQAと過去の検品結果から合格・修正継続・停止を判定（編集なし） |
| `add_item` | `voice` | セリフ1件追加（実音声長を返す） |
| `add_script` | —— | 複数セリフ一括追加（**実音声長で重なり自動回避**） |
| `edit_item` | `property` | `item_id`（推奨）またはframe+layerで変更。`expected_revision`対応 |
| `edit_item` | `keyframe` | Animationキーフレームの set/remove/clear。`at`は相対フレーム |
| `get_info` | `keyframes` | 指定アイテムのキーフレーム一覧 |
| `edit_item` | `delete` | `item_id`（推奨）または位置/レイヤーで削除。`expected_revision`対応 |
| `edit_item` | `move` | ファイル名指定でアイテムを移動 |
| `edit_item` | `select` | `item_id`（推奨）、frame、layerのいずれかを明示して選択（clearで全解除） |
| `edit_item` | `resolve_overlaps` | 重なり解消（gap指定可） |
| `edit_item` | `shift` | from_frame以降をdeltaフレーム一括シフト |
| `get_info` | `selection` | 選択中アイテムの詳細取得 |
| `get_info` | `commands` | 利用可能コマンド一覧 |
| `get_info` | `effects` | 指定アイテムのエフェクト現在値取得 |
| `control` | `undo` / `redo` | 元に戻す / やり直し |
| `control` | `split` / `align` | 再生位置で分割 / 整列 |

`get_info/project` の `isSaved` と `hasUnsavedChanges` は真偽値です。YMM4から保存状態を取得できない場合は両方とも `null` とし、未保存ではないと推測しません。

編集対象は明示してください。`property` と `keyframe` は `item_id` または `frame` と `layer` の両方が必要です。`face_param` と `effect` は `frame` と `layer` の両方が必要です。`select` は `item_id`、`frame`、`layer` のいずれか、または `clear=true` が必要です。対象を省略しても先頭アイテムや全アイテムを暗黙に選ぶことはありません。

`edit_item/property` は設定後に同じプロパティを読み直し、要求値と一致したときだけ `verified: true` を返します。YMM4側で値が丸められた、または拒否された場合は `PROPERTY_VERIFY_FAILED` と実際の値を返します。setterの例外や読み取り失敗は `outcome_unknown: true` です。失敗後は再送前に `items` で現在状態を確認してください。

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

**構造化タイムラインQA：**

`action="validate"` はタイムラインを変更せず、同一レイヤー内の重複・内部の空白・プロジェクト尺超過・期待アイテムの有無を検査します。
期待値には `item_id` と `revision` を指定できるため、移動後のアイテムや取得後に変更されたアイテムも正確に照合できます。

```python
action="validate",
duration=1800,
include_gaps=False,  # 空白を意図したレイヤーでは警告を省略
subtitle_layers=[9],  # レイヤー9のTextItemを字幕としてセリフとの対応を検査（省略可）
expected=[
    {"item_id": "native:...", "revision": "..."},
]
```

結果は `passed`、0〜100の `score`、エラー・警告件数の `summary`、および `issues` を返します。
各Issueには可能な範囲で `frame_range`、`item_ids`、`suggested_fix` が含まれます。`problems` は既存クライアント互換のための `issues` の別名です。
先頭フレームより前の空白は警告せず、アイテム同士の内部空白だけを `GAP` として報告します。
`subtitle_layers` を指定すると、各 VoiceItem の発話と時間が重なる指定レイヤーの TextItem を探し、空白・改行を除いた本文が一致しなければ `SUBTITLE_MISSING` または `SUBTITLE_TEXT_MISMATCH` を返します。テロップ用レイヤーは指定しないでください。発話テキストを取得できない場合は `SUBTITLE_CHECK_SKIPPED` を警告します。字幕が発話の全時間を覆うかどうかは検査しません。
映像の見切れや音量などはこの検査の対象外なので、`preview` / `watch` と組み合わせて確認してください。

**修正ループの停止判定：** `action="qa_gate"` は同じ検査条件で現在の `validate` を実行し、`qa_history`（過去の `validate` 結果を古い順に並べた配列）と比較します。`decision` は `pass` / `repair` / `stop`、`reason_code` は `QA_PASSED` / `QA_ISSUES_REMAIN` / `QA_REGRESSED` / `QA_STALLED` / `REPAIR_LIMIT_REACHED` / `TIME_LIMIT_REACHED` / `API_LIMIT_REACHED` です。結果の `qa` を次回の `qa_history` に追加してください。既定では修正3回が上限で、同じ問題群が2回連続した場合も停止します。経過時間とAPI回数は呼び出し側が `elapsed_seconds` / `api_calls` を数え、必要に応じて `max_seconds` / `max_api_calls` を指定します。品質悪化時の `suggested_action: consider_checkpoint_rollback` は提案のみで、ロールバックは自動実行されません。映像・音声の品質判定は行いません。

各 `validate` 結果の `criteria_hash` は期待アイテム・尺・空白・字幕レイヤーの検査条件を表します。`qa_gate` は現在と履歴の条件が異なる場合、品質の変化を誤判定しないよう入力エラーを返します。

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

**HTTPの応答待ち時間：**
- `watch` / `record` は、実効録音時間（秒）+ 15秒を確保します。8秒録音なら23秒、30秒録音なら45秒です。
- 録音時間はプラグイン側で `watch`: 1000〜30000ms、`record`: 500〜30000ms に制限されます。待ち時間もこの範囲に合わせます。
- `seek_capture` は描画待ちを考慮して15秒、通常の情報取得・操作は10秒、クリップ書き出しは600秒です。
- 接続待ち時間は10秒のままです。接続失敗時はプラグイン起動の案内、処理待ちのタイムアウト時は時間超過の案内を返します。
- タイムアウト後もYMM4側の処理が続く可能性があるため、自動再試行はしません。処理完了・状態を確認してから次の操作を行ってください。
- HTTP接続はMCPサーバー内で再利用し、終了時に解放します。

**開発者向け回帰テスト（YMM4不要）：**

リポジトリのルートで実行します。MCP SDKは既存のハンドラAPIに対応する1系を使用します（`requirements.txt` で2系を除外）。

```bash
python -m pip install -r mcp-server/requirements.txt
python -m unittest discover -s mcp-server -p 'test_*.py' -v
```

`test_jobs.py` は書き出しパス検証・成果物ヘッダ検査・ジョブdispatchを、YMM4なしで確認します。`test_editplan.py` はEditPlanの検証・差分・冪等再送・シーン単位rollback・checkpoint dispatchを確認します。HTTP通信をモックし、待ち時間、エラー分類、接続再利用・終了処理を確認します。実際の録音・画像保存やYMM4の起動は行いません。

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

`inspect` の `target` と `path` は URL クエリ値としてエンコードされます。失敗時は `success: false`、`error_code`、`error`、`retryable: false`、`outcome_unknown` を返します。主なコードは `INVALID_ARGUMENT`、`TARGET_NOT_FOUND`、`PATH_NOT_FOUND`、`MEMBER_NOT_FOUND`、`MEMBER_READ_ONLY`、`METHOD_NOT_FOUND`、`COMMAND_NOT_FOUND`、`COMMAND_UNAVAILABLE`、`REFLECTION_SET_FAILED`、`REFLECTION_INVOKE_FAILED`、`COMMAND_FAILED` です。実行中の例外で結果が確定できない場合は `outcome_unknown: true` になります。

`set` は設定後の値を読み戻し、一致した場合だけ `verified: true` を返します。値が丸められるなどして一致しない場合や読み戻せない場合は `REFLECTION_VERIFY_FAILED` です。`outcome_unknown: true` のときは自動再送せず、`get` または `inspect` で現在値を確認してください。

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

> **モデルの指定について**:
> 既定のモデルは `gemini-2.5-flash` です。高精度が必要な場合は、ツール呼び出し時に `model="gemini-2.5-pro"` を指定できます。
> また、`mcp-server` フォルダ内に `config.json` を作成し、以下のように記述することでデフォルトのモデルを変更することも可能です。
> ```json
> {
>   "default_model": "gemini-2.5-pro"
> }
> ```
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
| YMM4バージョン | v4.47以降（.NET 10対応）が必要 |
| watchの上限 | 画像+音声データが大きいため最大8秒程度推奨 |
| 初回キャプチャ | シーク直後の1枚目は黒になることがある（描画待ち）。正常動作 |
| ポート競合 | 8765が競合する場合はツール画面で停止し、ポートを変更・保存してから再起動 |
| 内部API | `IMainViewModel` の実装はYMM4バージョンにより異なる場合あり |
