# 完成動画のデコード検証（Issue #61）

MCPの `get_info/decoded_export_qa` は、Python MCPサーバーが動いているPC上の動画ファイルを FFmpeg で実際にデコードします。従来の `get_info/export_qa`（YMM4ホスト側のMP4コンテナ/メタデータ検証）を置き換えるものではなく、追加の検品です。

## 前提

- Windows等のMCPサーバー実行環境に `ffmpeg` と `ffprobe` をインストールし、両方を `PATH` から実行できるようにする。
- `path` は**Python MCPサーバーのPCからアクセスできる**絶対パスを指定する。YMM4を別PCで動かしている場合、そのPCのパスを直接指定しても読み取れない。
- 書き出しジョブが完了してから実行する。

## MCPの呼び出し例

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","require_audio":true,"sample_seconds":8}
```

`qa_gate` に `decoded_export_check` として同じ設定を渡すと、構造QAと併せて合否判定に反映できます。

```json
{"action":"qa_gate","decoded_export_check":{"path":"C:/Videos/final.mp4","require_audio":true,"sample_seconds":8}}
```

検証は映像・音声の最初の8秒と最後の8秒（尺が短い場合は冒頭だけ）を別々にデコードし、失敗箇所・ストリーム種別を `issues` に返します。音声が不要な作品は `require_audio:false` を指定できます。サンプル長は1〜15秒です。FFmpegがない、タイムアウトした、メタデータが不正な場合は**成功扱いにしません**。

**重要な制約:** `fully_decoded:false` のとおり、途中の未サンプル区間、黒画面、長時間無音、音割れ、主観的な映像品質は検査しません。実機での再生確認や、別途 `visual_qa` / `audio_qa` を併用してください。動画全体の完全なデコード検証や映像・音声品質判定は今後の課題です。
