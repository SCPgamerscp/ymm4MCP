# YMM4 4.56.1.1 キーフレーム実機スモークテスト（Issue #140）

このテストは **YMM4が起動しているWindows PCでのみ** 実行します。CIのモックテストだけでは元の `KEYFRAME_METHOD_UNAVAILABLE` 不具合の解消を保証できません。

## 準備

1. **必ずプロジェクトを別名でコピー・バックアップ**し、コピー側をYMM4で開いてください。保存していない作業のあるプロジェクトには使用しないでください。
2. コピー側に画像アイテムを1つ置き、YMM4の `GET /api/items` またはMCPの `get_info/items` から `item_id` と長さを取得します。
3. 対象画像の `Zoom` と `X` の **両方でまだ使用していない、0より大きくアイテム尺未満の相対フレーム** を選びます。
4. YMM4 MCPプラグインを起動し、`mcp-server` の依存関係をインストールしたPythonを使用します。接続情報は `YMM4_CONNECTION_FILE` または通常の `%LOCALAPPDATA%\YMM4MCP\connection.json` から自動で読みます。トークンをログに貼らないでください。

PowerShellで（値・ID・フレームはテスト用の実データに置き換えます）:

```powershell
python mcp-server/check_keyframe_compat.py --item-id "native:YOUR-TEST-IMAGE" --at 30 --zoom 1.25 --x 80 --confirm-test-project
```

**スクリプトの動作:** 編集前に対象画像の**全Animationプロパティ**（Zoom/X以外も含む）のキーフレーム一覧を取得して保全し、ZoomとXが利用可能で指定フレームに既存キーがないことを確認します。その後、順番に追加→API読戻しで位置と値を検証→追加したキーを削除し、**全プロパティのキーフレームが元どおりか**確認します。変更を伴うため、`--confirm-test-project` がないと実行できません。プロジェクトは自動保存しません。失敗時も安全と確認できる範囲で削除を試みます。ただし、書き込みが失敗した／応答が失われた場合に他の編集者による変更や共有Animationへの影響が見つかったときは、**他人のキーフレームを消さないため自動削除を拒否**します。削除APIの失敗や変更元を特定できない場合は **CLEANUP FAILED** と表示するため、バックアップから復元してください。

## 受け入れ条件の残り（手動確認）

スモークテストがPASSした後も、同じコピー側プロジェクトでZoom/Xのアニメーションを実際に設定し、**プレビューの移動・ズーム、保存→YMM4で再読込後のキー維持、再編集・削除**を確認してください。これらがすべて成功し、関連PRがmainにマージされた場合にのみ #140 を完了扱いにします。

`python -m unittest discover -s mcp-server -p 'test_check_keyframe_compat.py' -v` はHTTPモックによるプロトコルテストです。YMM4実機は使用しません。


## 保存→再読込の読み取り専用検証（新規・任意）

従来の `check_keyframe_compat.py` は追加→読戻し→削除までで、**保存後にキーが残るかは確認しません**。次の二段階検証は **YMM4にGETだけを送り、編集・保存・再読込は自動実行しません**。

1. バックアップ済みの**使い捨て** `.ymmp` をYMM4で開き、画像アイテムのZoomとXに同じ相対フレーム（例: 30）でキーを追加します。**まだ保存しない**でください。 `GET /api/items` の `identity_persistent=true` な `item_id` が必要です。
2. 編集が未保存の状態（`GET /api/project` の `isSaved=false`）で、次を実行します。

```powershell
python mcp-server/check_keyframe_persistence.py capture --snapshot "$env:TEMP\\ymm4-keyframe-capture.json" --item-id "native:YOUR-TEST-IMAGE" --at 30 --zoom 1.25 --x 80 --confirm-test-project
```

3. **YMM4のUIから**使い捨てプロジェクトを保存し、いったん閉じて同じ `.ymmp` を開き直します。再読込はAPIで証明できないため、操作者が確認してください。
4. 同じPCで次を実行します。

```powershell
python mcp-server/check_keyframe_persistence.py verify --snapshot "$env:TEMP\\ymm4-keyframe-capture.json" --confirm-test-project --confirm-reopened
```

`verify` はプロジェクトパスの一致・保存済み状態・ファイル更新時刻の前進・画像の永続ID・Zoom/Xの期待値・**全Animationのキーフレーム一覧**を比較します。スナップショットはローカルのプロジェクトパスとキー値を含むため公開しないでください。既存スナップショットを上書きせず、接続トークンも保存しません。

**限界:** APIだけではYMM4のUIで実際に閉じて開き直した事実や、プレビュー表示の正しさは検証できません。テストのPASSだけでIssue #140をクローズせず、プレビュー・再編集・削除の実機確認とPRマージも必要です。CIの `test_check_keyframe_persistence.py` はHTTPモックであり、YMM4実機を使用しません。
