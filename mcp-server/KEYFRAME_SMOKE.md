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

**スクリプトの動作:** ZoomとXが利用可能で指定フレームに既存キーがないことを確認し、順番に追加→API読戻しで位置と値を検証→追加したキーを削除して読戻します。変更を伴うため、`--confirm-test-project` がないと実行できません。プロジェクトは自動保存しません。失敗時も削除を試みますが、YMM4が削除APIを拒否する場合は **CLEANUP FAILED** と表示し、バックアップから復元する必要があります。

## 受け入れ条件の残り（手動確認）

スモークテストがPASSした後も、同じコピー側プロジェクトでZoom/Xのアニメーションを実際に設定し、**プレビューの移動・ズーム、保存→YMM4で再読込後のキー維持、再編集・削除**を確認してください。これらがすべて成功し、関連PRがmainにマージされた場合にのみ #140 を完了扱いにします。

`python -m unittest discover -s mcp-server -p 'test_check_keyframe_compat.py' -v` はHTTPモックによるプロトコルテストです。YMM4実機は使用しません。
