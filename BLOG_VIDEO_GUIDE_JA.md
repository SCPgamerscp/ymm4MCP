# ymm4MCP 徹底解説ガイド（動画・ブログ向け）

このドキュメントは **「ymm4MCPって何？どう動く？何がすごい？」** を、
解説動画や技術ブログでそのまま使えるように整理したものです。

---

## 1. 3行でわかる ymm4MCP

1. **YMM4（ゆっくりMovieMaker4）をAIから直接操作する仕組み**です。  
2. C#プラグインがYMM4内でHTTP APIを公開し、PythonのMCPサーバーがそのAPIをMCPツールとしてClaudeへ橋渡しします。  
3. AIは「台本作成 → セリフ配置 → プレビュー確認（画像/音声） → 修正」まで一連の編集ループを回せます。

---

## 2. 全体アーキテクチャ（役割分担）

```text
[Claude]
   │ (MCP/stdio)
[Python MCP Server: mcp-server/server.py]
   │ (HTTP localhost:8765)
[YMM4 C# Plugin: YMM4McpPlugin/McpHttpServer.cs]
   │ (YMM4内部API/リフレクション)
[YMM4本体]
```

### それぞれの責務

- **Claude（MCPクライアント）**
  - `ymm4_interact` と `ymm4_preview` という2つの道具を呼ぶ。  
- **Python MCPサーバー**
  - MCPツール定義を持ち、引数を受けてYMM4プラグインのHTTP APIに変換。  
  - 画像は `ImageContent`、情報は `TextContent` で返す。  
- **YMM4 C#プラグイン**
  - `HttpListener` で `/api/...` エンドポイントを公開。  
  - YMM4のViewModel/Timelineにアクセスし、実際の編集操作を実行。  

---

## 3. このプロジェクトの強み

### 3-1. 「AIが見て・聞いて・直す」閉ループを持っている

単なる自動台本生成ではなく、以下を連続実行できます。

1. 映像キャプチャで現状確認
2. 音声録音で発話/無音を確認（RMS付き）
3. セリフやプロパティを編集
4. 再確認

この「観測→編集→検証」があるため、**動画制作の実運用に近い**です。

### 3-2. add_scriptで一括投入しやすい

`add_script` は文字数から長さ（フレーム）を推定し、
`start_frame` から連続配置できるため、
長尺のゆっくり台本投入が高速です。

### 3-3. スキルで動画ジャンル別の作法を分離

`mcp-server/skills/` に実況・解説・茶番・ストーリーの設計指針を分けており、
制作ルールをプロンプトに毎回書かなくてよい構成になっています。

---

## 4. 主要コンポーネント解説

## 4-1. YMM4プラグイン層（C#）

### エントリーポイント
- `McpToolPlugin` がYMM4ツールとして「MCP連携サーバー」画面を追加。

### UI/状態管理
- `McpViewModel` が Start/Stop/ClearLog を管理。
- `McpView.xaml` から起動状態・URL・ログを見える化。

### HTTP実行基盤
- `McpHttpServer` が `http://localhost:8765/` を待ち受け。
- `HandleRequest` のルーティングで `/api/...` を処理。
- API実行時は `Application.Current.Dispatcher.Invoke` を用い、
  UIスレッド上でYMM4オブジェクトを安全に操作。

### 実装上のポイント
- YMM4内部オブジェクトは公開APIが限定されるため、
  **リフレクション**でプロパティ/コマンドへアクセス。
- そのため `/api/debug/...` 系エンドポイントが多く、
  バージョン差異の調査・追従に使える設計。

---

## 4-2. MCPサーバー層（Python）

### ツール構成
- `ymm4_interact`：情報取得、再生制御、セリフ追加、プロパティ編集など。
- `ymm4_preview`：画像キャプチャ、シーク付きキャプチャ、音声録音、watch。

### 返却データの作り
- 通常情報はJSONテキスト化して返す。
- プレビュー画像は `ImageContent` で返し、LLMがそのまま視覚入力として扱える。
- `record` / `watch` ではWAVをローカル保存し、ファイルパスとRMSを返す。

### add_scriptの仕様要点
- 引数：`start_frame`, `fps`, `chars_per_sec`, `lines[]`
- 行ごとに文字数から長さを推定（最低1秒）
- `total_frames` を返すため次シーン接続がしやすい

---

## 5. APIの使い分け（制作目線）

## 5-1. 制作の最小セット

- 情報確認：`/api/status`, `/api/project`, `/api/items`
- 台本投入：`/api/items/voice` or MCP `add_script`
- 編集：`/api/items/prop`, `/api/items/delete`
- プレビュー：`/api/preview/seek`, `/api/preview/watch`
- 保存：`/api/project/save`

## 5-2. 実運用で効く流れ

1. `watch` で素材のシーン感を把握
2. セクションごとに `add_script`
3. 100f刻み等で `seek_capture`
4. 被り・尺ズレを `edit_item`
5. 最終 `watch` で同期確認
6. `save`

---

## 6. 解説動画・ブログ向けの語り口テンプレ

## 6-1. 冒頭30秒テンプレ（動画用）

> 「このymm4MCPは、ゆっくりMovieMaker4をAIから直接操作するための仕組みです。  
> ClaudeがMCPでPythonサーバーに命令し、PythonサーバーがYMM4プラグインのHTTP APIを叩きます。  
> つまり“AIが台本を書く”だけでなく、“実際にタイムラインへ配置して、映像と音声を確認し、修正まで回せる”のが特徴です。」

## 6-2. 比較軸（導入メリット）

- 従来：台本生成と編集確認が分断される
- ymm4MCP：編集対象（YMM4）にAIが直接作用する
- 結果：試作速度が上がり、修正ループが短くなる

---

## 7. 注意点・制約（正直に伝えると信頼される）

- YMM4の内部構造依存があるため、バージョン差で一部操作が壊れる可能性。
- `watch` は画像+音声を同時に扱うため、長時間取得は重くなりやすい。
- ローカルHTTP（8765番）を使うので、競合時はポート変更が必要。

---

## 8. このプロジェクトを紹介するときの着地

おすすめの締め方：

> 「ymm4MCPは“AIに動画制作を丸投げする”というより、  
> “AIをYMM4の共同編集者にする”ための基盤です。  
> 台本生成・配置・確認・修正のループがつながることで、  
> ゆっくり動画制作の速度と再現性を大きく上げられます。」

