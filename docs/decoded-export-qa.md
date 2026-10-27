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

期待する書き出し尺が分かっている場合は `expected_duration_seconds` を指定できます。実際の動画尺と期待尺の差が `duration_tolerance_seconds`（既定1秒、0〜60秒）を超えると、デコード自体が成功しても `EXPORT_DURATION_MISMATCH`（error）として不合格になります。期待尺は有限の正数で最大7日です。省略時は従来どおり尺の期待値チェックを行いません。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","expected_duration_seconds":600,"duration_tolerance_seconds":0.5}
```

期待する動画形式を指定できます。 `expected_width` / `expected_height`（各1〜16384）、`expected_fps`（0超〜240、`fps_tolerance` は既定0.1）、`expected_audio_sample_rate`（8000〜384000Hz）、`expected_audio_channels`（1〜8）を指定すると、**全ての該当ストリーム**を検証します。不一致は `EXPORT_VIDEO_FORMAT_MISMATCH` / `EXPORT_AUDIO_FORMAT_MISMATCH`、必要なメタデータ欠落は `EXPORT_VIDEO_FORMAT_UNKNOWN` / `EXPORT_AUDIO_FORMAT_UNKNOWN`（いずれもerror）です。音声形式を指定すると `require_audio:false` でも音声が必須になります。さらに `expected_video_streams`（1〜16、カバー画像を除く）と `expected_audio_streams`（0〜16）で映像・音声ストリーム数を検証でき、期待値と異なる場合は `EXPORT_VIDEO_STREAM_COUNT_MISMATCH` / `EXPORT_AUDIO_STREAM_COUNT_MISMATCH`（error）を返します。音声ストリーム数0を期待する場合は `require_audio:false` を指定してください。未指定なら従来の検証動作です。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","expected_width":1920,"expected_height":1080,"expected_fps":30,"fps_tolerance":0.1,"expected_audio_sample_rate":48000,"expected_audio_channels":2}
```

`qa_gate` に `decoded_export_check` として同じ設定を渡すと、構造QAと併せて合否判定に反映できます。

```json
{"action":"qa_gate","decoded_export_check":{"path":"C:/Videos/final.mp4","require_audio":true,"sample_seconds":8}}
```

## オプション：全トラックの末尾までの継続確認

音声トラックはBGMなどの演出で動画より早く終了することがあるため、`EXPORT_STREAM_ENDS_EARLY` は既定で**警告**です。音声を含めた全ストリームが動画の末尾まで続くことを完成条件とする場合、`require_tracks_to_end:true` を指定してください。映像・音声の各トラックがコンテナ尺より**1秒超**短いと `EXPORT_STREAM_ENDS_EARLY`（error）になり、必要なストリーム時刻メタデータを取得できない場合は `EXPORT_STREAM_TAIL_UNVERIFIABLE`（error）になります。短いBGM・SEなどを別トラックとして含む作品では指定しないでください。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","require_audio":true,"require_tracks_to_end":true}
```

`qa_gate` の `decoded_export_check` でも同じフラグを使えます。これはFFprobeの各トラック時刻に基づく検証で、映像・音声の内容そのものが意図どおりかは判定しません。

検証は映像・音声の冒頭・中間（長尺の場合）・末尾を別々にデコードし、失敗箇所・ストリーム種別を `issues` に返します。音声が不要な作品は `require_audio:false` を指定できます。サンプル長は1〜15秒です。FFmpegがない、タイムアウトした、メタデータが不正な場合は**成功扱いにしません**。

サンプル内の **1秒以上の黒画面候補**（`blackdetect`）と **2秒以上の長い無音候補**（`silencedetect`）も検出し、`black_frames` / `long_silences` と `issues` に `severity: "warning"` で返します。これらは意図的なフェードや演出の可能性があるため、警告だけでは `passed:false` にしません。各イベントの `sample_start_seconds` はサンプル位置、`filter_start_seconds` / `filter_end_seconds` はFFmpegフィルタが報告したタイムスタンプです。動画全体の時刻と混同しないでください。

音声サンプルでは FFmpeg の `astats` から最大サンプルピーク（dBFS）を読み取り、`audio_peaks` に記録します。ピークが **-0.1 dBFS 以上**なら `EXPORT_SAMPLE_AUDIO_PEAK_NEAR_CLIPPING` を警告として返します。これは音割れの可能性を示すヒントであり、クリッピングの確定診断ではありません（True Peak / LUFS の測定ではありません）。警告のみで `passed:false` にはしません。

## オプション：全尺デコード検証

`full_decode:true` を指定すると、冒頭・中間（長尺の場合）・末尾サンプルQAに加え、**動画内のすべての映像ストリームと存在する音声ストリームを、先頭から末尾までFFmpegでデコード**します。サンプルだけでは見逃す中間区間のデコード破損も検出できます。既定値は `false` なので既存の動作・処理時間は変わりません。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","full_decode":true,"require_audio":true}
```

`qa_gate` からも `decoded_export_check:{"path":"C:/Videos/final.mp4","full_decode":true}` として利用できます。成功した場合は `verification_scope:"full_audio_video_decode"` と `fully_decoded:true` が返ります。デコードエラーは `EXPORT_FULL_DECODE_FAILED`、時間切れは `EXPORT_FULL_DECODE_TIMEOUT` として不合格になります。

処理が長時間化するため、**15分（900秒）を超える動画では全尺デコードを実行せず** `EXPORT_FULL_DECODE_DURATION_LIMIT` を返します。各ストリームの実行時間には60〜900秒の上限を設けています。動画尺が15分以下でも遅いPCや重いコーデックでは時間切れになることがあります。全尺デコードはCPU負荷が高く、バックグラウンドジョブではありません。

**重要な制約:** 全尺デコードはコーデックの破損を調べる機能であり、映像・音声の主観的品質や全区間の黒画面・無音・音割れを保証しません。`black_frames` / `long_silences` / `audio_peaks` の検査範囲は引き続き**冒頭・中間（長尺の場合）・末尾のサンプルのみ**（`signal_scope:"sampled_only"`）です。意図的な暗転や無音は警告です。YMM4実機での再生確認と併用してください。
