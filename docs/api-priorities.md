# 自律編集 API の現状と実装優先順位

Issue #17 の完了条件である API 仕様の整理と優先順位を記録する。ここで「実装済み」は API の存在を示す。YMM4 の全バージョンで動作することや、エンドツーエンドの制作が完了することを意味しない。

| #17 の項目 | 現行の入口 | 残る境界 |
| --- | --- | --- |
| 事後検証 | `get_info/items`、`get_info/keyframes`、`validate` | 変更したフィールドの期待値を1回で照合する操作はない |
| 構造の整合性 | `validate`、`qa_gate` | 構造以外の検査は指定したプレビュー位置やWAVに限る |
| チェックポイント | `control/checkpoint`、`control/rollback` | rollback は追加アイテムの削除のみ。削除・プロパティ・キーフレーム変更の復元は `.ymmp` バックアップが必要 |
| 構造化エラー | 多くの編集 API に `error_code` | 一部の旧 Reflection API は同じ契約になっていない |
| 入力検証 | MCP の schema と個別の範囲検証 | 実ホストのプロパティ範囲は動的に検出する必要がある |
| dry-run | `add_script`、`plan_edit`、`duck_bgm` | 全ての低レベル操作に共通する dry-run はない |
| テロップ・字幕 | `add_item/text`、`edit_item/property` | 書式プリセットと画面内配置の検査が不足 |
| 画像・立ち絵 | `add_item/image`、`add_item/tachie`、`add_item/face` | 素材差し替え、表情・ポーズの意味的な指定が不足 |
| BGM・SE | `get_info/assets`、`add_item/audio`、`duck_bgm` | タグ検索、音量正規化が不足 |
| 最終書き出し | `control/export`、`get_info/job`、`get_info/export_qa` | 実フレームのデコード検証は行わない |

## P0: 安全な切り取り・復元

1. `split_item(item_id, at, expected_revision)` は相対フレームがアイテム内にあることを検証し、左右の `item_id` と `revision` を返す。UI選択状態に依存しないこと。
2. `trim_item(item_id, start_offset, end_offset, expected_revision)` は元素材へのオフセットと長さを一緒に更新し、処理後のソース範囲を返す。未対応アイテムは変更前に `TRIM_UNSUPPORTED` を返す。
3. `ripple_delete(start_frame, end_frame, layers, expected_revisions)` は影響するアイテムを先に列挙し、全件の編集可能性と競合を検証する。変更できないアイテムがあれば何も変更せず `RIPPLE_CONFLICT` を返す。成功時は新しい位置とIDの一覧を返す。
4. トランザクションの復元範囲を add/delete/property/keyframe/move/split に拡張し、失敗時は変更前状態と照合する。完全復元を保証できないホストでは操作を拒否し、バックアップの場所を返す。

## P1: 意図と検査結果を直接結び付ける

- `verify_items(expected=[{item_id, revision?, frame?, layer?, length?, text?}])` は現在値との不一致を `issues` に返す。読取専用で、曖昧な ID は成功にしない。
- `apply_expression(item_id, expression, reaction?, at?, expected_revision)` は登録済みのキャラクターと利用可能な表情・ポーズの一覧を先に照合する。対応がない場合は `EXPRESSION_UNAVAILABLE` を返し、勝手に別の表情へ置き換えない。
- `qa_gate` の視聴覚検査は元のフレーム範囲、検出元、重大度、対象ID、修正候補を構造化して返す。サンプリングの間を検査済みと表示しない。

## P2: 運用と素材発見

- 素材検索はフォルダを明示した読み取り専用操作から始め、重複ハッシュ、使用中の item ID、欠損を返す。タグ付けは別の永続ストアとして設計する。
- スタイルプリセットは version、fps、layer、字幕と音量の既定値を持ち、適用前に `plan_edit` と同じ差分を提示する。

共通契約: 変更対象には `item_id` と `expected_revision` を使い、複数操作では `idempotency_key` を受ける。`success=false` と `error_code`、部分成功時の `outcome_unknown` とバックアップ場所を明示する。実装の着手順は P0 の小さな操作と復元検証、P1 の事後検証、P2 の素材・プリセットとする。
