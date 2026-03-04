---
name: ymm4-mcp
description: |
  YukkuriMovieMaker4（YMM4）をMCPで操作するスキル。
  ゆっくり実況動画の制作補助に使う。
  以下の場面で必ず使うこと：
  - 「ゆっくり実況を作って」「セリフを追加して」「タイムラインを編集して」
  - 「映像を確認して」「音声を確認して」「BGMを調べて」
  - 「セリフと映像を合わせて」「重複をチェックして」
  - 「スケルトンキング」「マインクラフト」等の動画内容への言及
  - YMM4・ゆっくりムービーメーカー・ゆっくり動画に関する操作全般
---

# YMM4 MCP スキル

YukkuriMovieMaker4をMCP経由で操作し、ゆっくり実況動画を自動生成・編集するスキル。

---

## 前提条件

- YMM4が起動済み
- YMM4McpPluginが有効（ツール → MCP連携サーバー → 起動）
- MCPサーバーがport 8765で動作中

---

## APIエンドポイント一覧

### プロジェクト・アイテム系
```
GET  /api/status
GET  /api/project
GET  /api/items
POST /api/project/save
POST /api/timeline/duration
```

### セリフ・スクリプト系
```
POST /api/voice/add
POST /api/script/add
POST /api/item/edit
POST /api/item/delete
```

### 映像確認系
```
GET  /api/preview/capture
POST /api/preview/seek
GET  /api/preview/position
POST /api/playback/play
POST /api/playback/stop
```

### 音声系
```
POST /api/preview/record
POST /api/preview/watch    ← 映像＋音声同時取得
```

---

## MCPツールの使い方

### ymm4_interact（操作系）

```python
# アイテム一覧取得
action="get_info", sub_action="items"

# セリフ一括追加（文字数から長さ自動計算）
action="add_script",
start_frame=0, fps=30, chars_per_sec=5,
lines=[
    {"layer": 7, "character": "ゆっくり霊夢",   "text": "セリフ内容"},
    {"layer": 8, "character": "ゆっくり魔理沙", "text": "セリフ内容"},
]
# ※ add_scriptは前のセリフ終了直後に次を配置する
#   シーンごとにstart_frameを指定して複数回呼ぶこと

# アイテム削除
action="edit_item", sub_action="delete", layer=7, frame=0

# プロパティ変更（フレーム移動など）
action="edit_item", sub_action="property",
layer=8, frame=186, prop="Frame", value="210"
```

### ymm4_preview（映像・音声確認系）

```python
# 特定フレームに移動して画像取得
action="seek_capture", frame=738

# 映像＋音声を同時取得（最も強力）
action="watch",
frame=0,
duration_ms=5000,          # 最大8000ms推奨
capture_interval_ms=2000   # 2000ms以上推奨（小さいと1MBエラー）

# 音声のみ録音
action="record", duration_ms=5000
```

---

## 標準ワークフロー：ゆっくり実況を作る

### Step 1：映像全体を把握する

watchで複数シーンを確認する。

```python
# 大まかな流れを掴む（例）
for frame in [0, 500, 1500, 3000, 5000]:
    action="watch", frame=frame,
    duration_ms=4000, capture_interval_ms=2000
```

### Step 2：シーン分けと台本作成

映像の内容をもとに台本を作る。

**フレーム計算の目安（5文字/秒・30fps）:**

| 文字数 | フレーム | 秒数 |
|---|---|---|
| 20文字 | 120f | 4秒 |
| 30文字 | 180f | 6秒 |
| 40文字 | 240f | 8秒 |
| 50文字 | 300f | 10秒 |

**被り防止のルール：**
- 前のセリフ終了f = 開始f + length
- 次のセリフ開始f ≥ 前のセリフ終了f + 10f（余裕を持たせる）
- 同レイヤー内で被りがないか必ず確認

### Step 3：シーンごとにadd_scriptで配置

```python
# 例：0fから配置
action="add_script", start_frame=0, lines=[
    {"layer": 7, "character": "ゆっくり霊夢",   "text": "冒頭のセリフ"},
    {"layer": 8, "character": "ゆっくり魔理沙", "text": "続きのセリフ"},
]
# → 戻り値のtotal_framesが次のstart_frameの目安

# 次のシーン（例：460fから）
action="add_script", start_frame=460, lines=[...]
```

### Step 4：100fごとに映像確認

```python
for frame in range(0, 最終フレーム, 100):
    action="seek_capture", frame=frame
    # 確認：セリフと映像が合っているか
```

**確認項目：**
- セリフのテキストが映像の内容と合っているか
- 表示フレームのシーンが正しいか
- 前後のセリフが重複していないか

### Step 5：ズレ・重複を修正

```python
# 被りチェック（アイテム一覧から手計算）
# 同レイヤーで: item[n].frame < item[n-1].frame + item[n-1].length → 被り！

# フレーム位置を修正
action="edit_item", sub_action="property",
layer=8, frame=186, prop="Frame", value="210"

# セリフを削除して作り直す場合
action="edit_item", sub_action="delete", layer=7, frame=954
```

---

## 音声・BGM分析

### watchで録音してPythonで解析

```python
# 1. watchで録音（capture_interval_msを大きくして画像を減らす）
action="watch", frame=0, duration_ms=8000, capture_interval_ms=8000
# → C:\...\ymm4プラグイン\watch_result.wav に保存される

# 2. Pythonで解析（WAVはfmt=3 IEEE float 32bit, 48kHz, 2ch）
import struct, math
with open(path, "rb") as f: data = f.read()
# dataチャンクを探してサンプル取得
# left = samples[::2]  ← 左チャンネルのみ
```

**分析指標：**

| 指標 | 意味 |
|---|---|
| RMS > 0.01 | 音声あり |
| ZCR 高い | 高音・速いリズム |
| 主要周波数 200-800Hz | ピアノ・ストリングス系 |
| 0.5秒ごとRMSの変動 | 曲のリズム感 |

---

## よくある問題と対処

| 問題 | 原因 | 対処 |
|---|---|---|
| watchでToo largeエラー | データが1MB超 | duration_ms↓、capture_interval_ms↑ |
| キャプチャが真っ黒（1枚目） | シーク直後の描画待ち | 正常。2枚目以降を使う |
| セリフの被り | add_scriptの自動配置がズレた | frame指定を手動で調整する |
| BGMしか録音できない | ゆっくり音声が再生されていない | YMM4の再生が開始されているか確認 |
| MCP接続エラー | サーバーが未起動 | YMM4でツール→MCP連携サーバー→起動 |

---

## レイヤー構成の慣例

```
Layer 0:   動画ファイル（VideoItem）
Layer 1-6: エフェクト・字幕等
Layer 7:   ゆっくり霊夢のセリフ
Layer 8:   ゆっくり魔理沙のセリフ
Layer 9+:  追加キャラ
```

---

## キャラクター設定

| キャラ | 一人称 | 備考 |
|---|---|---|
| ゆっくり霊夢 | 私（わたし） | ツッコミ役 |
| ゆっくり魔理沙 | 私（わたし） | ボケ役。**「俺」は使わない** |

> ⚠️ **魔理沙の一人称は「私」。「俺」「僕」は誤り。セリフ生成時は必ず「私」を使うこと。**
