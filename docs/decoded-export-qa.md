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

## オプション：コーデック指定・黒画面/無音/フリーズの合否ゲート

`expected_video_codec` / `expected_audio_codec` を指定すると、全ての映像/音声トラックのFFprobe `codec_name` を照合します。例: `h264`, `hevc`, `aac`, `opus`。未取得なら `EXPORT_VIDEO_CODEC_UNKNOWN` / `EXPORT_AUDIO_CODEC_UNKNOWN`、不一致なら `EXPORT_VIDEO_CODEC_MISMATCH` / `EXPORT_AUDIO_CODEC_MISMATCH`（いずれもerror）です。指定する名前は小文字英数字・アンダースコアのみ（1〜32文字）です。音声コーデックを指定すると音声トラックは必須になります。

`max_black_seconds` / `max_silence_seconds` / `max_freeze_seconds` を指定すると、**実際に観測した**各黒画面・無音・静止区間が閾値を超えた場合、`EXPORT_BLACK_DURATION_EXCEEDED` / `EXPORT_SILENCE_DURATION_EXCEEDED` / `EXPORT_FREEZE_DURATION_EXCEEDED`（error）として不合格にします。閾値は0〜900秒です。未指定時は従来どおり警告のみです。複数のサンプル窓で重複観測した同一区間は結合後に判定します。**サンプル検証だけでは未観測区間の品質を保証できません**。全尺の品質条件を検証したい場合は `full_decode:true` と `full_signal_scan:true` を併用してください。

さらに `expected_video_frames`（1〜1,000,000,000）を指定すると、全映像トラックのFFprobe `nb_frames` と照合します。フレーム数不明なら `EXPORT_VIDEO_FRAME_COUNT_UNKNOWN`、不一致なら `EXPORT_VIDEO_FRAME_COUNT_MISMATCH`（error）です。VFRやコンテナによっては `nb_frames` が得られないため、必須の場合のみ指定してください。

`min_audio_peak_dbfs` / `max_audio_peak_dbfs`（各 -120〜12dBFS）を指定すると、**各サンプル**の最大サンプルピークが下限未満・上限超過の場合に `EXPORT_SAMPLE_AUDIO_PEAK_TOO_LOW` / `EXPORT_SAMPLE_AUDIO_PEAK_TOO_HIGH`（error）にします。ピークが測定不能な場合は `EXPORT_SAMPLE_AUDIO_PEAK_UNVERIFIABLE`（error）で、完全無音の `-inf` は下限を下回るものとして扱います。ピーク値はLUFS/True Peakではなく、音量感や歪みの確定診断ではありません。

`max_stereo_imbalance_db`（0〜120dB）は、ステレオ音声の左右サンプルピーク差が上限を超えた場合に `EXPORT_SAMPLE_STEREO_IMBALANCE_EXCEEDED`（error）にします。左右の測定値が揃わない場合は `EXPORT_SAMPLE_STEREO_METRICS_UNAVAILABLE`（error）です。チャンネル数が不明・ステレオ以外の場合は `EXPORT_SAMPLE_STEREO_CHANNELS_UNVERIFIABLE`（error）です。音声品質ゲートを指定すると `require_audio:false` でも音声ストリームが必須になります。片側が完全無音で他方に音がある場合も不合格になります。

`require_audio_peak_metrics:true` は、デコードに成功した音声サンプルでもFFmpeg `astats` の有効なピーク測定が存在しない場合に `EXPORT_SAMPLE_AUDIO_METRICS_UNAVAILABLE`（error）を返します。完全無音を示す `-inf` は有効な測定値です。既定は `false` で互換性を維持します。すべて `qa_gate.decoded_export_check` でも指定できます。

```json
{"action":"qa_gate","decoded_export_check":{"path":"C:/Videos/final.mp4","expected_video_codec":"h264","expected_audio_codec":"aac","max_black_seconds":2,"max_silence_seconds":3,"max_freeze_seconds":4,"require_audio_peak_metrics":true,"full_decode":true,"full_signal_scan":true}}
```

## オプション：全音声トラックの統合ラウドネス（LUFS）検証

`min_integrated_lufs` / `max_integrated_lufs`（-70〜10 LUFS）を指定すると、FFmpeg `loudnorm` の**正規化前入力** `input_i` を、音声トラックごとに**全尺**で測定します。単一サンプルのピークdBFSでは分からない、完成動画全体の音量感の過小・過大を検出するための品質ゲートです。

`full_decode:true` が必須です。各トラックの追加フルパスが必要なので、全尺デコードと共有する15分の実行予算内で動作します。予算超過、FFmpegエラー、測定JSON欠落、`-inf` など測定不能な値は**不合格**にし、音量が適切だったと誤判定しません。音声トラックがない場合も不合格です。既定では測定を行わず、既存のQA挙動は変わりません。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","full_decode":true,"min_integrated_lufs":-24,"max_integrated_lufs":-12}
```

結果の `integrated_loudness` は `[{"stream_ordinal":0,"integrated_lufs":-18.5}]` の形式です。閾値違反は `EXPORT_LOUDNESS_TOO_QUIET` / `EXPORT_LOUDNESS_TOO_LOUD`、測定失敗は `EXPORT_LOUDNESS_FAILED` / `EXPORT_LOUDNESS_UNVERIFIABLE` / `EXPORT_LOUDNESS_TIMEOUT`（すべてerror）です。`qa_gate.decoded_export_check` からも指定できます。LUFSは映像全体の音量基準であり、**セリフとBGMを別々に識別した相対音量評価やTrue Peakの保証ではありません**。

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

## オプション：全尺の黒画面・フリーズ・無音検出

`full_decode:true` と `full_signal_scan:true` を**同時に指定**すると、通常の全尺デコードに加えて全映像・音声ストリームをフィルタで走査し、サンプル窓の間にある黒画面・フリーズ・長時間無音も検出します。

```json
{"action":"get_info","sub_action":"decoded_export_qa","path":"C:/Videos/final.mp4","full_decode":true,"full_signal_scan":true}
```

`qa_gate.decoded_export_check` にも同じ2フラグを指定できます。成功時は `signal_scope:"full"` となり、`black_frames` / `frozen_frames` / `long_silences` に動画全尺の検出候補を返します。これらは演出として意図的な場合があるため**警告**であり、自動的に不合格にはしません。ストリームの開始時刻がずれる場合も動画時刻へ補正します。サンプル由来の重複検出区間は統合します。

全尺信号走査も15分の全尺デコード予算を共有し、フィルタ失敗・タイムアウト・不正な診断は `EXPORT_FULL_SIGNAL_FAILED` / `EXPORT_FULL_SIGNAL_TIMEOUT` / `EXPORT_FULL_SIGNAL_TIMESTAMP_INVALID` で不合格になります。`full_signal_scan` 単独指定は拒否します。処理時間とCPU負荷が増えるため、既定値は `false` です。音量ピーク・左右チャンネルの `astats` は引き続きサンプル区間のみです。

**重要な制約:** 通常の `full_decode:true` だけでは黒画面・無音・フリーズの検出はサンプル範囲のみ（`signal_scope:"sampled_only"`）です。全尺信号走査を指定しても、字幕の見切れ・意図的な演出かどうか・True Peak / LUFS・主観的品質は判定できません。YMM4実機での再生確認と併用してください。
