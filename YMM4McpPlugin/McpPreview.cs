using System;
using System.Globalization;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        // ■ プレビューキャプチャ
        // ================================================================

        /// <summary>現在のプレビュー画面をPNG(base64)で返す（GDI PrintWindow方式でDirectX対応）</summary>
        private static object CapturePreview(HttpListenerRequest req)
        {
            try
            {
                string? b64 = null;
                string? err = null;
                int captureW = 0, captureH = 0;

                Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        var window = Application.Current.MainWindow;
                        if (window == null) { err = "MainWindow not found"; return; }

                        // WindowsFormsHost（実映像エリア）→ PreviewView の順で探す
                        UIElement? previewElem = FindVisualByName(window, "WindowsFormsHost");
                        if (previewElem == null) previewElem = FindVisualByName(window, "PreviewView");
                        if (previewElem == null)
                            foreach (var n in new[] { "Preview", "PreviewArea", "PlayerView", "VideoPreview" })
                            { previewElem = FindVisualByName(window, n); if (previewElem != null) break; }

                        var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                        GetWindowRect(hwnd, out RECT winRect);

                        int cropLeft = 0, cropTop = 0;
                        int cropW = winRect.Right - winRect.Left;
                        int cropH = winRect.Bottom - winRect.Top;

                        if (previewElem != null)
                        {
                            // PointToScreen でスクリーン座標を取得し、ウィンドウ左上との差分をピクセル単位で計算
                            var screenPt = previewElem.PointToScreen(new System.Windows.Point(0, 0));
                            cropLeft = (int)(screenPt.X - winRect.Left);
                            cropTop  = (int)(screenPt.Y - winRect.Top);
                            if (previewElem is FrameworkElement fe)
                            {
                                var dpi = VisualTreeHelper.GetDpi(window);
                                cropW = (int)(fe.ActualWidth  * dpi.DpiScaleX);
                                cropH = (int)(fe.ActualHeight * dpi.DpiScaleY);
                            }
                        }

                        // ウィンドウ全体をPrintWindowでキャプチャ（DirectX/OpenGL対応）
                        int fullW = winRect.Right - winRect.Left;
                        int fullH = winRect.Bottom - winRect.Top;
                        if (fullW <= 0 || fullH <= 0) { err = "ウィンドウサイズ取得失敗"; return; }

                        using var bmp = new Bitmap(fullW, fullH, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                        using (var g = Graphics.FromImage(bmp))
                        {
                            var hdc = g.GetHdc();
                            bool ok = PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT);
                            g.ReleaseHdc(hdc);
                            if (!ok) { err = "PrintWindow失敗"; return; }
                        }

                        // PreviewView の領域だけ切り出す
                        cropLeft = Math.Max(0, Math.Min(cropLeft, fullW - 1));
                        cropTop  = Math.Max(0, Math.Min(cropTop,  fullH - 1));
                        cropW    = Math.Max(1, Math.Min(cropW, fullW - cropLeft));
                        cropH    = Math.Max(1, Math.Min(cropH, fullH - cropTop));

                        using var cropped = new Bitmap(cropW, cropH);
                        using (var g2 = Graphics.FromImage(cropped))
                            g2.DrawImage(bmp, new System.Drawing.Rectangle(0, 0, cropW, cropH),
                                              new System.Drawing.Rectangle(cropLeft, cropTop, cropW, cropH),
                                              GraphicsUnit.Pixel);

                        using var ms = new MemoryStream();
                        cropped.Save(ms, ImageFormat.Png);
                        b64 = Convert.ToBase64String(ms.ToArray());
                        captureW = cropW; captureH = cropH;
                    }
                    catch (Exception ex) { err = ex.Message; }
                });

                if (err != null) return new { success = false, error = err };
                return (object)new { success = true, format = "png", width = captureW, height = captureH, element = "PreviewView(GDI)", image = b64 };
            }
            catch (Exception ex) { return new { success = false, error = ex.Message }; }
        }

        /// <summary>指定フレームにシークしてからキャプチャ</summary>
        private static async Task<object> SeekAndCapture(HttpListenerRequest req)
        {
            try
            {
                var body = await ReadBody(req);
                int frame = GetInt(body, "frame", 0);

                // PreviewViewModel.SeekAsync(Int32) を呼ぶ
                var preview = Application.Current.Dispatcher.Invoke(() => GetPreviewViewModel());
                if (preview == null) return new { success = false, error = "PreviewViewModel not found" };

                try
                {
                    await InvokeAsyncMethod(preview, "SeekAsync", frame);
                }
                catch (Exception ex)
                {
                    return (object)new { success = false, error = $"SeekAsync failed: {ex.Message}" };
                }

                await Task.Delay(300);
                return CapturePreview(req);
            }
            catch (Exception ex) { if (ex is ArgumentException or JsonException) throw; return (object)new { success = false, error = ex.Message }; }
        }

        /// <summary>現在の再生位置(フレーム)を返す</summary>
        private static object GetPlaybackPosition()
        {
            try
            {
                return Application.Current.Dispatcher.Invoke(() =>
                {
                    var preview = GetPreviewViewModel();
                    if (preview == null) return (object)new { success = false, error = "PreviewViewModel not found" };

                    // PreviewViewModelのプロパティからフレーム・合計フレームを取得
                    object? frame = GetPropValue(preview, "CurrentFrame")
                                 ?? GetPropValue(preview, "Frame")
                                 ?? GetPropValue(preview, "Position");
                    object? totalFrames = GetPropValue(preview, "TotalFrames")
                                      ?? GetPropValue(preview, "Duration");

                    return (object)new { success = true, currentFrame = frame, totalFrames };
                });
            }
            catch (Exception ex) { return new { success = false, error = ex.Message }; }
        }

        /// <summary>システム音声をWASAPIループバックで録音してbase64 WAVで返す</summary>
        private static async Task<object> RecordAudio(HttpListenerRequest req)
        {
            try
            {
                var body = await ReadBody(req);
                int durationMs = GetInt(body, "duration_ms", 3000);
                durationMs = Math.Clamp(durationMs, 500, 30000);

                using var capture = new NAudio.Wave.WasapiLoopbackCapture();
                var waveFormat = capture.WaveFormat;
                var buffer = new OrderedAudioChunks();
                var tcs = new TaskCompletionSource<bool>();

                capture.DataAvailable += (s, e) =>
                {
                    if (e.BytesRecorded > 0)
                    {
                        buffer.Add(e.Buffer, e.BytesRecorded);
                    }
                };
                capture.RecordingStopped += (s, e) => tcs.TrySetResult(true);

                try
                {
                    capture.StartRecording();
                    await Task.Delay(durationMs);
                    try { capture.StopRecording(); } catch { }

                    // RecordingStopped が発火するまで最大2秒待つ
                    await Task.WhenAny(tcs.Task, Task.Delay(2000));

                    // バッファを結合
                    var allBytes = buffer.ToArray();

                    // WAVヘッダーを付けてbase64化
                    using var ms = new MemoryStream();
                    using (var writer = new NAudio.Wave.WaveFileWriter(ms, waveFormat))
                        writer.Write(allBytes, 0, allBytes.Length);

                    var b64 = Convert.ToBase64String(ms.ToArray());
                    double rms = CalcRms(allBytes, waveFormat.BitsPerSample);

                    return (object)new
                    {
                        success = true,
                        duration_ms = durationMs,
                        sample_rate = waveFormat.SampleRate,
                        channels = waveFormat.Channels,
                        bits = waveFormat.BitsPerSample,
                        bytes_recorded = allBytes.Length,
                        rms_level = Math.Round(rms, 4),
                        has_audio = rms > 0.0005,
                        format = "wav",
                        audio = b64
                    };
                }
                finally
                {
                    try { capture.StopRecording(); } catch { }
                }
            }
            catch (Exception ex) { if (ex is ArgumentException or JsonException) throw; return (object)new { success = false, error = ex.Message }; }
        }

        /// <summary>
        /// 指定フレームから再生しながら、音声録音＋一定間隔で映像キャプチャを同時実行して返す
        /// POST body: { "frame": 300, "duration_ms": 5000, "capture_interval_ms": 1000 }
        /// </summary>
        private static async Task<object> WatchScene(HttpListenerRequest req)
        {
            try
            {
                var body = await ReadBody(req);
                int startFrame = GetInt(body, "frame", 0);
                int durationMs = Math.Clamp(GetInt(body, "duration_ms", 5000), 1000, 30000);
                int intervalMs = Math.Clamp(GetInt(body, "capture_interval_ms", 1000), 500, 10000);

                // 1) シーク
                var preview = Application.Current.Dispatcher.Invoke(() => GetPreviewViewModel());
                if (preview == null) return new { success = false, error = "PreviewViewModel not found" };
                await RunOnUi(() => InvokeAsyncMethod(preview, "SeekAsync", startFrame));
                await Task.Delay(300);

                // 2) 録音・再生・キャプチャを並行実行
                using var capture = new NAudio.Wave.WasapiLoopbackCapture();
                var waveFormat = capture.WaveFormat;
                var audioBuffer = new OrderedAudioChunks();
                var recordTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                capture.DataAvailable += (s, e) =>
                {
                    if (e.BytesRecorded > 0)
                    {
                        audioBuffer.Add(e.Buffer, e.BytesRecorded);
                    }
                };
                capture.RecordingStopped += (s, e) => recordTcs.TrySetResult(true);
                try
                {
                    capture.StartRecording();

                    await InvokeAsyncMethod(preview, "TogglePlayAsync");

                    var frames = new System.Collections.Generic.List<object>();
                    var captureTask = Task.Run(async () =>
                    {
                        int elapsed = 0;
                        while (elapsed < durationMs)
                        {
                            await Task.Delay(intervalMs);
                            elapsed += intervalMs;
                            string? b64 = null; int fw = 0, fh = 0;
                            Application.Current.Dispatcher.Invoke(() =>
                            { var r = CaptureCurrentFrame(); b64 = r.b64; fw = r.w; fh = r.h; });
                            if (b64 != null)
                                frames.Add(new { time_ms = elapsed, width = fw, height = fh, image = b64 });
                        }
                    });

                    await Task.Delay(durationMs);
                    try { await InvokeAsyncMethod(preview, "StopAsync"); } catch { }
                    try { capture.StopRecording(); } catch { }
                    await Task.WhenAny(recordTcs.Task, Task.Delay(2000));
                    await captureTask;

                    var allBytes = audioBuffer.ToArray();
                    using var ms = new MemoryStream();
                    using (var writer = new NAudio.Wave.WaveFileWriter(ms, waveFormat))
                        writer.Write(allBytes, 0, allBytes.Length);
                    var audioB64 = Convert.ToBase64String(ms.ToArray());
                    double rms = CalcRms(allBytes, waveFormat.BitsPerSample);

                    return (object)new
                    {
                        success = true,
                        start_frame = startFrame,
                        duration_ms = durationMs,
                        audio = new
                        {
                            format = "wav",
                            sample_rate = waveFormat.SampleRate,
                            channels = waveFormat.Channels,
                            bits = waveFormat.BitsPerSample,
                            rms_level = Math.Round(rms, 4),
                            has_audio = rms > 0.0005,
                            data = audioB64
                        },
                        frames
                    };
                }
                finally
                {
                    try { await InvokeAsyncMethod(preview, "StopAsync"); } catch { }
                    try { capture.StopRecording(); } catch { }
                }
            }
            catch (Exception ex) { if (ex is ArgumentException or JsonException) throw; return (object)new { success = false, error = ex.Message }; }
        }

        /// <summary>プロジェクトのFPS(と解像度)を取得する。タイムスタンプ→フレーム変換の基準として使う。</summary>
        private object GetProjectFps()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var video = ReadProjectVideoInfo(vm);
                return (object)new { success = true, fps = video.fps ?? 30, video.width, video.height,
                    detected = video.fps.HasValue,
                    note = video.fps == null ? "FPS自動検出失敗のため30を仮定" : null };
            });
        }

        private static (int? fps, int? width, int? height) ReadProjectVideoInfo(object vm)
        {
            object? project = GetPropObj(vm, "Project") ?? GetPropObj(vm, "project") ?? GetPropObj(vm, "_project");
            var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
            int? fps = null, width = null, height = null;
            foreach (var src in new object?[] { project, project == null ? null : GetPropObj(project, "VideoInfo"), tvm,
                tvm != null ? GetType_Field(tvm, "scene") : null,
                tvm != null ? GetType_Field(tvm, "timeline") : null })
            {
                if (src == null) continue;
                foreach (var name in new[] { "FPS", "Fps", "FrameRate", "VideoFPS", "VideoInfo" })
                {
                    var v = GetNestedInt(src, name);
                    if (v is > 0 and <= 240) fps ??= v;
                }
                foreach (var name in new[] { "Width", "VideoWidth", "ScreenWidth" })
                {
                    var v = GetNestedInt(src, name);
                    if (v is > 0 and <= 16384) width ??= v;
                }
                foreach (var name in new[] { "Height", "VideoHeight", "ScreenHeight" })
                {
                    var v = GetNestedInt(src, name);
                    if (v is > 0 and <= 16384) height ??= v;
                }
            }
            return (fps, width, height);
        }

        private static object? GetType_Field(object o, string name)
        {
            try { return o.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(o); }
            catch { return null; }
        }

        /// <summary>オブジェクトのプロパティ/フィールド(ReactiveProperty含む)を辿って int を取り出す。VideoInfo等のネストも1段掘る。</summary>
        private static int? GetNestedInt(object o, string name)
        {
            try
            {
                var val = GetPropValue(o, name) ?? GetType_Field(o, name);
                if (val == null) return null;
                if (val is int i) return i;
                if (int.TryParse(val.ToString(), out int p)) return p;
                // ネストオブジェクトなら FPS/Width 等をもう1段
                foreach (var inner in new[] { "FPS", "Fps", "FrameRate", "Width", "Height", "Value" })
                {
                    var iv = GetPropValue(val, inner);
                    if (iv != null && int.TryParse(iv.ToString(), out int p2)) return p2;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 【案A：プレビュー解析用】指定区間をフレーム単位でシーク＆キャプチャし、連番PNGをディスクに保存する。
        /// 同時に区間の音声をWAVで録音して同じフォルダに保存する。
        /// 戻り値は保存先フォルダと各フレームのメタ情報(frame番号/timeMs/ファイル名)。
        /// Python側はこのフォルダの画像群＋WAVをGeminiへ渡して解析する。
        ///
        /// POST body: {
        ///   "startFrame": 0, "endFrame": 300,   // 解析するフレーム範囲
        ///   "stepFrames": 3,                    // 何フレームおきにキャプチャするか(粒度)
        ///   "outputDir": "C:\\...\\clip",       // 省略時は一時フォルダ
        ///   "recordAudio": true                 // 区間音声も録音するか
        /// }
        /// </summary>
        private async Task<object> ExportClip(HttpListenerRequest req)
        {
            NAudio.Wave.WasapiLoopbackCapture? capture = null;
            TaskCompletionSource<bool>? recordTcs = null;
            object? preview = null;
            try
            {
                var body = await ReadBody(req);
                int startFrame = GetInt(body, "startFrame", 0);
                int endFrame = GetInt(body, "endFrame", startFrame + 300);
                int stepFrames = Math.Max(1, GetInt(body, "stepFrames", 3));
                bool recordAudio = !(body.TryGetValue("recordAudio", out var ra) && ra.ValueKind == JsonValueKind.False);
                string outputDir = GetStr(body, "outputDir", "");

                if (endFrame < startFrame) return new { success = false, error = "endFrame は startFrame 以上である必要があります" };
                // 無制限のフレーム数で固まらないよう上限を設ける。int 加算だと MaxValue 近辺でオーバーフローして無限ループになる。
                long totalShots = ((long)endFrame - startFrame) / stepFrames + 1;
                if (totalShots > 2000) return new { success = false, error = $"フレーム数が多すぎます({totalShots})。stepFramesを大きくするか範囲を狭めてください" };

                if (string.IsNullOrEmpty(outputDir))
                    outputDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ymm4_clip_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                System.IO.Directory.CreateDirectory(outputDir);

                preview = Application.Current.Dispatcher.Invoke(() => GetPreviewViewModel());
                if (preview == null) return new { success = false, error = "PreviewViewModel not found" };

                // FPS取得(タイムスタンプ計算用)
                int fps = 30;
                try { var f = GetProjectFps(); var fp = f.GetType().GetProperty("fps")?.GetValue(f); if (fp != null) int.TryParse(fp.ToString(), out fps); } catch { }
                if (fps <= 0) fps = 30;

                OrderedAudioChunks? audioBuffer = null;
                NAudio.Wave.WaveFormat? waveFormat = null;
                string? wavPath = null;

                var savedFrames = new List<object>();
                var swTotal = System.Diagnostics.Stopwatch.StartNew();

                if (recordAudio)
                {
                    await InvokeAsyncMethod(preview, "SeekAsync", startFrame);
                    await Task.Delay(200);
                    capture = new NAudio.Wave.WasapiLoopbackCapture();
                    waveFormat = capture.WaveFormat;
                    audioBuffer = new OrderedAudioChunks();
                    recordTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    capture.DataAvailable += (s, e) =>
                    {
                        if (e.BytesRecorded > 0)
                        {
                            audioBuffer.Add(e.Buffer, e.BytesRecorded);
                        }
                    };
                    capture.RecordingStopped += (s, e) => recordTcs.TrySetResult(true);
                    capture.StartRecording();
                    await InvokeAsyncMethod(preview, "TogglePlayAsync");
                }

                int index = 0;
                for (long frame64 = startFrame; frame64 <= endFrame; frame64 += stepFrames)
                {
                    int frame = (int)frame64;
                    // 再生中(録音中)はSeekせず再生位置のキャプチャを行うとズレるため、
                    // 録音時は実時間ベースで待機しながらキャプチャ、非録音時はSeekしてキャプチャ
                    if (recordAudio)
                    {
                        int targetMs = (int)((frame - startFrame) * 1000.0 / fps);
                        int waitMs = targetMs - (int)swTotal.ElapsedMilliseconds;
                        if (waitMs > 0) await Task.Delay(waitMs);
                    }
                    else
                    {
                        await InvokeAsyncMethod(preview, "SeekAsync", frame);
                        await Task.Delay(120); // 描画待ち
                    }

                    string? b64 = null; int w = 0, h = 0;
                    Application.Current.Dispatcher.Invoke(() => { var r = CaptureCurrentFrame(); b64 = r.b64; w = r.w; h = r.h; });
                    if (b64 == null) continue;

                    string fileName = $"frame_{index:D5}_f{frame}.png";
                    string filePath = System.IO.Path.Combine(outputDir, fileName);
                    try { System.IO.File.WriteAllBytes(filePath, Convert.FromBase64String(b64)); }
                    catch (Exception ex) { savedFrames.Add(new { frame, error = ex.Message }); continue; }

                    double timeMs = (frame - startFrame) * 1000.0 / fps;
                    savedFrames.Add(new { index, frame, timeMs = Math.Round(timeMs, 1), file = fileName, width = w, height = h });
                    index++;
                }

                double rms = 0; bool hasAudio = false;
                if (recordAudio && capture != null && waveFormat != null && audioBuffer != null && recordTcs != null)
                {
                    try { await InvokeAsyncMethod(preview, "StopAsync"); } catch { }
                    try { capture.StopRecording(); } catch { }
                    await Task.WhenAny(recordTcs.Task, Task.Delay(2000));
                    var allBytes = audioBuffer.ToArray();
                    wavPath = System.IO.Path.Combine(outputDir, "audio.wav");
                    using (var fs = new FileStream(wavPath, FileMode.Create))
                    using (var writer = new NAudio.Wave.WaveFileWriter(fs, waveFormat))
                        writer.Write(allBytes, 0, allBytes.Length);
                    rms = CalcRms(allBytes, waveFormat.BitsPerSample);
                    hasAudio = rms > 0.0005;
                }

                return new
                {
                    success = true,
                    outputDir,
                    fps,
                    startFrame,
                    endFrame,
                    stepFrames,
                    frameCount = index,
                    durationMs = Math.Round((endFrame - startFrame) * 1000.0 / fps, 1),
                    audioFile = wavPath,
                    audioRms = Math.Round(rms, 4),
                    hasAudio,
                    frames = savedFrames
                };
            }
            catch (Exception ex) { if (ex is ArgumentException or JsonException) throw; return new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            finally
            {
                if (capture != null)
                {
                    if (preview != null)
                    {
                        try { await InvokeAsyncMethod(preview, "StopAsync"); } catch { }
                    }
                    try { capture.StopRecording(); } catch { }
                    if (recordTcs != null)
                    {
                        try { await Task.WhenAny(recordTcs.Task, Task.Delay(2000)); } catch { }
                    }
                    try { capture.Dispose(); } catch { }
                }
            }
        }

        /// <summary>現在フレームをキャプチャしてbase64を返す（Dispatcher内から呼ぶ）</summary>
        private static (string? b64, int w, int h) CaptureCurrentFrame()
        {
            try
            {
                var window = Application.Current.MainWindow;
                if (window == null) return (null, 0, 0);
                UIElement? previewElem = FindVisualByName(window, "WindowsFormsHost")
                                     ?? FindVisualByName(window, "PreviewView");
                var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                GetWindowRect(hwnd, out RECT winRect);
                int fullW = winRect.Right - winRect.Left, fullH = winRect.Bottom - winRect.Top;
                if (fullW <= 0 || fullH <= 0) return (null, 0, 0);
                int cropLeft = 0, cropTop = 0, cropW = fullW, cropH = fullH;
                if (previewElem != null)
                {
                    var screenPt = previewElem.PointToScreen(new System.Windows.Point(0, 0));
                    cropLeft = (int)(screenPt.X - winRect.Left);
                    cropTop  = (int)(screenPt.Y - winRect.Top);
                    if (previewElem is FrameworkElement fe)
                    {
                        var dpi = VisualTreeHelper.GetDpi(window);
                        cropW = (int)(fe.ActualWidth  * dpi.DpiScaleX);
                        cropH = (int)(fe.ActualHeight * dpi.DpiScaleY);
                    }
                }
                using var bmp = new Bitmap(fullW, fullH, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    var hdc = g.GetHdc();
                    PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT);
                    g.ReleaseHdc(hdc);
                }
                cropLeft = Math.Max(0, Math.Min(cropLeft, fullW - 1));
                cropTop  = Math.Max(0, Math.Min(cropTop,  fullH - 1));
                cropW    = Math.Max(1, Math.Min(cropW, fullW - cropLeft));
                cropH    = Math.Max(1, Math.Min(cropH, fullH - cropTop));
                using var cropped = new Bitmap(cropW, cropH);
                using (var g2 = Graphics.FromImage(cropped))
                    g2.DrawImage(bmp, new System.Drawing.Rectangle(0, 0, cropW, cropH),
                                      new System.Drawing.Rectangle(cropLeft, cropTop, cropW, cropH),
                                      GraphicsUnit.Pixel);
                using var outMs = new MemoryStream();
                cropped.Save(outMs, ImageFormat.Png);
                return (Convert.ToBase64String(outMs.ToArray()), cropW, cropH);
            }
            catch { return (null, 0, 0); }
        }

        private static double CalcRms(byte[] data, int bitsPerSample)
        {
            if (data.Length == 0) return 0;
            try
            {
                if (bitsPerSample == 32)
                {
                    // IEEE float 32bit
                    double sum = 0;
                    int count = data.Length / 4;
                    for (int i = 0; i < count; i++)
                    {
                        double s = BitConverter.ToSingle(data, i * 4);
                        sum += s * s;
                    }
                    return Math.Sqrt(sum / count);
                }
                if (bitsPerSample == 16)
                {
                    double sum = 0;
                    int count = data.Length / 2;
                    for (int i = 0; i < count; i++)
                    {
                        double s = BitConverter.ToInt16(data, i * 2) / 32768.0;
                        sum += s * s;
                    }
                    return Math.Sqrt(sum / count);
                }
                return 0;
            }
            catch { return 0; }
        }

    }
}
