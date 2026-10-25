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

サンプル内の **1秒以上の黒画面候補**（`blackdetect`）と **2秒以上の長い無音候補**（`silencedetect`）も検出し、`black_frames` / `long_silences` と `issues` に `severity: "warning"` で返します。これらは意図的なフェードや演出の可能性があるため、警告だけでは `passed:false` にしません。各イベントの `sample_start_seconds` はサンプル位置、`filter_start_seconds` / `filter_end_seconds` はFFmpegフィルタが報告したタイムスタンプです。動画全体の時刻と混同しないでください。

音声サンプルでは FFmpeg の `astats` から最大サンプルピーク（dBFS）を読み取り、`audio_peaks` に記録します。ピークが **-0.1 dBFS 以上**なら `EXPORT_SAMPLE_AUDIO_PEAK_NEAR_CLIPPING` を警告として返します。これは音割れの可能性を示すヒントであり、クリッピングの確定診断ではありません（True Peak / LUFS の測定ではありません）。警告のみで `passed:false` にはしません。

**重要な制約:** `fully_decoded:false` のとおり、途中の未サンプル区間、音割れの確定診断、主観的な映像品質は検査しません。黒画面・無音の検出も冒頭と末尾のサンプルだけであり、映像全体の品質を保証しません。実機での再生確認や、別途 `visual_qa` / `audio_qa` を併用してください。動画全体の完全なデコード検証やクリッピングの確定診断は今後の課題です。
