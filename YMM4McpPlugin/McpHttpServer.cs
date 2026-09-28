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
        // GDI/Win32 API（DirectX/OpenGLレンダリングのキャプチャ用）
        [DllImport("user32.dll")] static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
        [DllImport("user32.dll")] static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);
        [DllImport("user32.dll")] static extern IntPtr GetDesktopWindow();
        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int Left, Top, Right, Bottom; }
        const uint PW_RENDERFULLCONTENT = 2; // DirectX/OpenGL対応フラグ

        private readonly object _lifecycle = new();
        private HttpListener? _listener;
        private string _token = "";
        private bool _allowAdvanced;
        public McpSettings Settings { get; } = McpSettings.Load();
        public int Port => Settings.Port;
        public string BaseUrl => $"http://127.0.0.1:{Port}/";
        public bool IsRunning => _listener?.IsListening == true;
        public event Action<string>? LogMessage;
        public event Action? StateChanged;

        public void Start()
        {
            lock (_lifecycle)
            {
                if (IsRunning) return;
                Settings.Validate();
                McpSettings.PrepareDirectory();
                var listener = new HttpListener();
                listener.Prefixes.Add(BaseUrl);
                try
                {
                    listener.Start();
                    string token = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
                    string temp = McpSettings.ConnectionPath + ".tmp";
                    File.WriteAllText(temp, JsonSerializer.Serialize(new { api_base = BaseUrl + "api", token }));
                    File.Move(temp, McpSettings.ConnectionPath, true);
                    _token = token;
                    _allowAdvanced = Settings.AllowAdvanced;
                    _listener = listener;
                    _ = ListenLoop(listener);
                    RestorePersistedJobs();
                    RestoreEditBindings();
                    RestoreCheckpoints();
                }
                catch { listener.Close(); throw; }
            }
            StateChanged?.Invoke();
            Log($"MCPサーバー起動: {BaseUrl}");
        }

        public void Stop()
        {
            lock (_lifecycle)
            {
                var listener = _listener;
                _listener = null;
                listener?.Close();
                _token = "";
                CancelAndPersistJobs();
                // Keep the protected discovery file: clients can distinguish connection
                // refusal after shutdown; the token is rotated on the next start.
            }
            StateChanged?.Invoke();
            Log("MCPサーバー停止");
        }

        private async Task ListenLoop(HttpListener listener)
        {
            try
            {
                while (listener.IsListening)
                {
                    var ctx = await listener.GetContextAsync();
                    _ = HandleRequest(ctx);
                }
            }
            catch (HttpListenerException) { }
            catch (ObjectDisposedException) { }
            catch (Exception ex) { Log($"HTTP listener error: {ex.Message}"); }
            finally
            {
                lock (_lifecycle)
                {
                    listener.Close();
                    if (ReferenceEquals(_listener, listener)) _listener = null;
                }
                StateChanged?.Invoke();
            }
        }

        private bool IsAuthenticated(HttpListenerRequest request)
        {
            string expected = _token;
            string supplied = request.Headers["X-Ymm4-Token"] ?? "";
            return expected.Length > 0 && supplied.Length == expected.Length &&
                CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(expected), Encoding.UTF8.GetBytes(supplied));
        }

        private async Task HandleRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;
            bool editLock = false;
            try
            {
                string path = req.Url?.AbsolutePath ?? "/";
                if (!IsAuthenticated(req))
                {
                    await WriteJson(res, 401, new { success = false, error_code = "UNAUTHORIZED", error = "X-Ymm4-Token is required" });
                    return;
                }
                if (req.Headers["Origin"] != null || req.Headers["Sec-Fetch-Site"] != null)
                {
                    await WriteJson(res, 403, new { success = false, error_code = "BROWSER_REQUEST_DENIED", error = "Browser requests are not supported" });
                    return;
                }
                if (!_allowAdvanced && (path.StartsWith("/api/reflect/", StringComparison.Ordinal) || path.StartsWith("/api/debug/", StringComparison.Ordinal)))
                {
                    await WriteJson(res, 403, new { success = false, error_code = "ADVANCED_DISABLED", error = "Enable advanced APIs in the plugin settings and restart" });
                    return;
                }
                if (req.HttpMethod == "POST" && !IsNonBlockingPost(path))
                {
                    editLock = await _editGate.WaitAsync(0);
                    if (!editLock)
                    {
                        await WriteJson(res, 409, Failure("EDIT_BUSY", "別の操作が進行中です。完了後に状態を確認してください"));
                        return;
                    }
                }
                Log($"{req.HttpMethod} {path}");
                object? result = await TryRouteJobs(req, path);
                result ??= await TryRouteEdits(req, path);
                result ??= (req.HttpMethod, path) switch
                {
                    ("GET", "/api/status") => GetStatus(),
                    ("GET", "/api/capabilities") => GetCapabilities(),
                    ("GET", "/api/project") => GetProjectInfo(),
                    ("GET", "/api/items") => GetTimelineItems(),
                    ("GET", "/api/media/export-qa") => GetExportQa(req),
                    ("GET", "/api/media/info") => GetMediaFileInfo(req),
                    ("GET", "/api/media/assets") => GetMediaAssets(req),
                    ("GET", "/api/characters") => GetCharacters(),
                    ("GET", "/api/media/audio-qa") => GetAudioQa(req),
                    ("POST", "/api/items/video") => await AddNativeItem(req, "video"),
                    ("POST", "/api/items/audio") => await AddNativeItem(req, "audio"),
                    ("POST", "/api/items/image") => await AddNativeItem(req, "image"),
#if DEBUG
                    ("GET", "/api/debug/vm") => DebugViewModel(),
                    ("GET", "/api/debug/timeline") => DebugTimeline(),
                    ("GET", "/api/debug/items") => DebugItems(),
                    ("GET", "/api/debug/scene") => DebugScene(),
                    ("GET", "/api/debug/scenefields") => DebugSceneFields(),
                    ("GET", "/api/debug/toolbar") => DebugToolBar(),
                    ("GET", "/api/debug/itemtoolbar") => DebugItemToolBar(),
                    ("GET", "/api/debug/voicecmd") => DebugVoiceCmd(),
                    ("GET", "/api/debug/menuitem") => DebugMenuItem(),
                    ("GET", "/api/debug/props") => GetProps(req),
                    ("GET", "/api/debug/search") => SearchProps(req),
                    ("GET", "/api/debug/type") => GetTypeInfo(req),
                    ("GET", "/api/debug/timelinemethods") => DebugTimelineMethods(),
                    ("GET", "/api/debug/voicetypes") => DebugVoiceTypes(),
#endif
                    ("POST", "/api/items/text") => await AddNativeItem(req, "text"),
                    ("POST", "/api/items/voice") => await AddNativeItem(req, "voice"),
                    ("POST", "/api/items/reorder") => await ReorderItems(req),
                    ("POST", "/api/items/arrange") => await ArrangeItems(req),
                    ("POST", "/api/items/move") => await MoveItem(req),
                    ("POST", "/api/items/tachie") => await AddNativeItem(req, "tachie"),
                    ("POST", "/api/items/face") => await AddNativeItem(req, "face"),
                    ("POST", "/api/items/face/param") => await SetFaceParam(req),
                    ("POST", "/api/items/effect/video") => await AddVideoEffect(req),
                    ("POST", "/api/items/effect") => await AddEffectToItem(req),
                    ("POST", "/api/items/prop") => await SetItemProp(req),
                    ("GET",  "/api/items/keyframes") => GetItemKeyframes(req),
                    ("POST", "/api/items/keyframe") => await SetItemKeyframe(req),
                    ("GET", "/api/effects/list") => ListEffects(),
                    ("GET", "/api/effects/describe") => DescribeEffect(req),
                    ("POST", "/api/items/delete") => await DeleteItems(req),
#if DEBUG
                    ("GET", "/api/debug/tachie") => DebugTachie(),
                    ("GET", "/api/debug/tachie/props") => DebugTachieItemProps(),
                    ("GET", "/api/debug/facetypes") => DebugFaceTypes(),
                    ("GET", "/api/debug/voiceitem/props") => DebugVoiceItemProps(),
#endif
                    ("POST", "/api/items/effect/audio") => await AddAudioEffect(req),
#if DEBUG
                    ("GET", "/api/debug/visualtree") => DebugVisualTree(req),
                    ("GET", "/api/debug/player") => DebugPlayer(),
#endif
                    ("GET", "/api/preview/capture") => CapturePreview(req),
                    ("POST", "/api/preview/seek") => await SeekAndCapture(req),
                    ("GET", "/api/preview/position") => GetPlaybackPosition(),
                    ("POST", "/api/preview/record") => await RecordAudio(req),
                    ("POST", "/api/preview/watch") => await WatchScene(req),
                    ("POST", "/api/preview/export-clip") => await ExportClip(req),
                    ("GET",  "/api/project/fps") => GetProjectFps(),
                    ("POST", "/api/timeline/duration") => await SetTimelineDuration(req),
                    ("POST", "/api/playback/play") => await PlaybackControl("play"),
                    ("POST", "/api/playback/stop") => await PlaybackControl("stop"),
                    ("POST", "/api/project/save") => SaveCurrentProject(),
                    ("POST", "/api/project/open") => await OpenProject(req),
                    ("POST", "/api/project/save-as") => await SaveProjectAs(req),
                    // ── 全機能アクセス用 汎用API ──────────────────────
                    ("POST", "/api/command") => await ExecCommandApi(req),
                    ("GET",  "/api/commands") => ListCommands(),
                    ("POST", "/api/reflect/get") => await ReflectGet(req),
                    ("POST", "/api/reflect/set") => await ReflectSet(req),
                    ("POST", "/api/reflect/invoke") => await ReflectInvoke(req),
                    ("GET",  "/api/reflect/inspect") => InspectObject(req),
                    // ── タイムライン操作・情報取得 ─────────────────────
                    ("GET",  "/api/selection") => GetSelection(),
                    ("POST", "/api/items/select") => await SelectItems(req),
                    ("POST", "/api/timeline/resolve-overlaps") => await ResolveOverlaps(req),
                    ("POST", "/api/timeline/shift") => await ShiftItems(req),
                    ("GET",  "/api/items/effects") => GetItemEffects(req),
                    _ => null
                };
                if (result == null) { await WriteJson(res, 404, new { error = "Not Found", path }); return; }
                await WriteJson(res, 200, result);
            }
            catch (Exception ex)
            {
                Log($"エラー: {ex.Message}");
                try { await WriteJson(res, ex is ArgumentException || ex is JsonException ? 400 : 500,
                    new { success = false, error_code = ex is ArgumentException || ex is JsonException ? "INVALID_ARGUMENT" : "INTERNAL_ERROR", error = ex.Message, retryable = false }); }
                catch { res.Abort(); }
            }
            finally { if (editLock) _editGate.Release(); }
        }

        private static async Task WriteJson(HttpListenerResponse res, int status, object data)
        {
            res.StatusCode = status;
            res.ContentType = "application/json; charset=utf-8";
            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            var bytes = Encoding.UTF8.GetBytes(json);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes, 0, bytes.Length);
            res.Close();
        }

        // ── API ──────────────────────────────────────────────

        private object GetStatus() => new { status = "running", version = typeof(McpHttpServer).Assembly.GetName().Version?.ToString(3), port = Port, timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") };

        private object GetCapabilities() => new
        {
            success = true,
            api_schema_version = 6,
            plugin_version = typeof(McpHttpServer).Assembly.GetName().Version?.ToString(3),
            authentication = "X-Ymm4-Token",
            advanced_enabled = _allowAdvanced,
            item_types = new[] { "video", "audio", "image", "text", "voice", "tachie", "face" },
            features = new
            {
                media_import = true, character_discovery = true, script_dry_run_in_mcp = true,
                timeline_validation_in_mcp = true, stable_item_ids = true, optimistic_concurrency = true,
                persistent_item_ids = "native_when_available", final_video_export = true,
                resumable_jobs = true, transactions = true, scene_transactions = true,
                checkpoints = true, keyframe_api = true,
                automatic_visual_audio_qa = false, project_open_save_as = true,
                declarative_edits = true, idempotent_edits = true
            },
            limitations = new[] { "Native operations require an open YMM4 timeline and a compatible MainModel signature.",
                "Serialized API writes do not lock manual UI edits.", "A timed-out operation may continue; inspect state before retrying.",
                "Voice parameters are inherited from registered characters; per-request engine/style overrides are not implemented.",
                "Items without a native YMM4 identifier receive runtime-only IDs; check identity_persistent before storing an ID across restarts.",
                "Keyframe edits use the host Animation/KeyFrames API via reflection; inspect the item if KEYFRAME_METHOD_UNAVAILABLE is returned.",
                "Export and project open/save-as discover host methods at runtime. EXPORT_METHOD_UNAVAILABLE / EXPORT_DIALOG_REQUIRED / OPEN_METHOD_UNAVAILABLE mean this YMM4 build needs a path-taking API.",
                "Jobs continue after MCP disconnect but become interrupted when the YMM4 process exits. resume re-queues; it does not continue an in-progress encode.",
                "EditPlan apply treats each scene as a transaction: a failed scene deletes its own additions and keeps committed scenes. Extra timeline items are never deleted.",
                "Checkpoint rollback deletes items added after the snapshot. It does not recreate deleted items or restore property mutations. Use the optional .ymmp backup with project/open for a full restore." }
        };

        private object GetProjectInfo()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var model = GetMainModel(vm);
                string? path = GetPropValue(vm, "ProjectFilePath")?.ToString() ?? (model == null ? null : GetPropObj(model, "ProjectFilePath")?.ToString());
                object? savedValue = GetPropValue(vm, "IsSaved") ?? (model == null ? null : GetPropValue(model, "IsSaved"));
                bool? isSaved = savedValue is bool saved ? saved : null;
                var video = ReadProjectVideoInfo(vm);
                return new { success = true, vmType = vm.GetType().FullName,
                    projectName = string.IsNullOrEmpty(path) ? null : Path.GetFileNameWithoutExtension(path),
                    projectPath = path, isSaved, hasUnsavedChanges = isSaved.HasValue ? !isSaved.Value : (bool?)null,
                    fps = video.fps, width = video.width, height = video.height };
            });
        }

        private object GetTimelineItems()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗", items = Array.Empty<object>() };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗", items = Array.Empty<object>() };
                var rawItems = GetPropEnum(tvm, "Items");
                var items = new List<object>();
                if (rawItems != null)
                {
                    foreach (var iv in rawItems)
                    {
                        var info = ReadItemInfo(iv);
                        var identity = GetItemIdentity(info.item);
                        items.Add(new
                        {
                            item_id = identity.id,
                            revision = GetItemRevision(info.item),
                            identity_persistent = identity.persistent,
                            layer = info.layer,
                            frame = info.frame,
                            length = info.length,
                            endFrame = (long)info.frame + info.length,
                            type = info.type,
                            text = info.text,
                            source_path = GetPropObj(info.item, "FilePath")?.ToString(),
                        });
                    }
                }
                return (object)new { items, count = items.Count };
            });
        }

        private async Task<object> SetTimelineDuration(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            int frames = GetInt(body, "frames", 1200);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };

                // プライベートフィールド "scene" と "timeline" からDurationを設定
                var results = new List<object>();
                foreach (var fieldName in new[] { "scene", "timeline" })
                {
                    var field = tvm.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                    if (field == null) continue;
                    var obj = field.GetValue(tvm);
                    if (obj == null) continue;

                    foreach (var propName in new[] { "Duration", "TotalFrame", "Length", "FrameCount", "TotalLength" })
                    {
                        var prop = obj.GetType().GetProperty(propName);
                        if (prop == null || !prop.CanWrite) continue;
                        try
                        {
                            var val = prop.GetValue(obj);
                            var valueProp = val?.GetType().GetProperty("Value");
                            if (valueProp != null) valueProp.SetValue(val, frames);
                            else prop.SetValue(obj, frames);
                            results.Add(new { field = fieldName, prop = propName, success = true, frames });
                        }
                        catch (Exception ex) { results.Add(new { field = fieldName, prop = propName, success = false, error = ex.Message }); }
                    }
                }

                if (results.Count == 0)
                {
                    var sceneField = tvm.GetType().GetField("scene", BindingFlags.NonPublic | BindingFlags.Instance);
                    var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                    var sceneObj = sceneField?.GetValue(tvm);
                    var timelineObj = timelineField?.GetValue(tvm);
                    return (object)new
                    {
                        success = false,
                        error = "Durationプロパティが見つかりません",
                        sceneProps = sceneObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                        timelineProps = timelineObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    };
                }
                return (object)new { success = true, results };
            });
        }

        private async Task<object> MoveItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            // filename: ファイル名（部分一致）, frame: 新しい開始フレーム, length: 新しい長さ（省略可）
            string filename = GetStr(b, "filename", "");
            if (string.IsNullOrWhiteSpace(filename)) throw new ArgumentException("filename is required");
            int newFrame = TimelineInputValidation.AtLeast(GetInt(b, "frame", 0), 0, "frame");
            int newLength = b.ContainsKey("length")
                ? TimelineInputValidation.AtLeast(GetInt(b, "length", -1), 1, "length") : -1;

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var fp = GetPropObj(item, "FilePath")?.ToString() ?? "";
                    if (!fp.Contains(filename)) continue;

                    var frameProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    var lengthProp = item.GetType().GetProperty("Length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    int oldFrame = (int)(frameProp?.GetValue(item) ?? 0);
                    int oldLength = (int)(lengthProp?.GetValue(item) ?? 0);
                    TimelineInputValidation.CheckPlacement(newFrame, newLength > 0 ? newLength : oldLength);
                    try
                    {
                        frameProp?.SetValue(item, newFrame);
                        if (newLength > 0) lengthProp?.SetValue(item, newLength);
                        return (object)new { success = true, filename, oldFrame, oldLength, newFrame, newLength };
                    }
                    catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
                }
                return (object)new { success = false, error = $"ファイルが見つかりません: {filename}" };
            });
        }

        // ════════════════════════════════════════════════════════════
        // ■ タイムライン高度操作・状態取得
        // ════════════════════════════════════════════════════════════

        /// <summary>各 TimelineItem の Item から frame/layer/length/type/text を抽出する共通ヘルパー。</summary>
        private static (int frame, int layer, int length, string type, string? text, object item) ReadItemInfo(object iv)
        {
            var item = GetPropObj(iv, "Item") ?? iv;
            int frame = 0, layer = 0, length = 0;
            try { frame = Convert.ToInt32(GetPropObj(item, "Frame") ?? GetPropObj(iv, "Frame") ?? 0); } catch { }
            try { layer = Convert.ToInt32(GetPropObj(item, "Layer") ?? GetPropObj(iv, "Layer") ?? 0); } catch { }
            try { length = Convert.ToInt32(GetPropObj(item, "Length") ?? GetPropObj(iv, "Length") ?? 0); } catch { }
            var text = (GetPropObj(item, "Serif") ?? GetPropObj(item, "Text") ?? GetPropObj(item, "FilePath") ?? GetPropObj(item, "Name"))?.ToString();
            return (frame, layer, length, item.GetType().Name, text, item);
        }

        /// <summary>現在UI上で「選択中」のアイテムの詳細（座標・レイヤー・フレーム・長さ・型）を返す。</summary>
        private object GetSelection()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                var selected = new List<object>();
                foreach (var iv in rawItems)
                {
                    // TimelineItemViewModel の IsSelected を確認
                    bool isSel = false;
                    try
                    {
                        var sv = GetPropObj(iv, "IsSelected");
                        if (sv is bool bsv) isSel = bsv;
                        else if (sv != null) { var vp = sv.GetType().GetProperty("Value"); if (vp?.GetValue(sv) is bool b2) isSel = b2; }
                    }
                    catch { }
                    if (!isSel) continue;
                    var (frame, layer, length, type, text, item) = ReadItemInfo(iv);
                    var identity = GetItemIdentity(item);
                    selected.Add(new
                    {
                        item_id = identity.id,
                        revision = GetItemRevision(item),
                        identity_persistent = identity.persistent,
                        frame,
                        layer,
                        length,
                        endFrame = (long)frame + length,
                        type,
                        text
                    });
                }
                // 現在の再生位置も付加
                object? playhead = null;
                try { playhead = GetPlaybackPosition(); } catch { }
                return (object)new { success = true, selectedCount = selected.Count, items = selected, playback = playhead };
            });
        }

        /// <summary>frame+layer 指定、または全クリアでアイテムを選択状態にする。</summary>
        private async Task<object> SelectItems(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            TimelineInputValidation.RequireSelectionTarget(b);
            int targetFrame = GetInt(b, "frame", -1);
            int targetLayer = GetInt(b, "layer", -1);
            string itemId = GetStr(b, "item_id", "");
            bool clear = b.TryGetValue("clear", out var ce) && ce.ValueKind == JsonValueKind.True;

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                object? idTarget = null;
                if (!clear && itemId.Length > 0)
                {
                    idTarget = FindItemById(rawItems, itemId, out bool ambiguous);
                    if (ambiguous) return Failure("ITEM_ID_AMBIGUOUS", "item_id が複数のアイテムに一致しました");
                    if (idTarget == null) return Failure("ITEM_NOT_FOUND", "item_id に一致するアイテムがありません");
                }

                int changed = 0;
                foreach (var iv in rawItems)
                {
                    var (frame, layer, _, _, _, item) = ReadItemInfo(iv);
                    bool shouldSelect;
                    if (clear) shouldSelect = false;
                    else if (idTarget != null) shouldSelect = ReferenceEquals(item, idTarget);
                    else shouldSelect = (targetFrame < 0 || frame == targetFrame) && (targetLayer < 0 || layer == targetLayer);
                    // clear時は全解除、それ以外は一致するものを選択
                    if (clear || shouldSelect)
                    {
                        var sProp = iv.GetType().GetProperty("IsSelected", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        if (sProp == null) continue;
                        try
                        {
                            var sv = sProp.GetValue(iv);
                            var vp = sv?.GetType().GetProperty("Value");
                            if (vp != null && vp.CanWrite) { vp.SetValue(sv, shouldSelect); changed++; }
                            else if (sProp.CanWrite) { sProp.SetValue(iv, shouldSelect); changed++; }
                        }
                        catch { }
                    }
                }
                return (object)new { success = true, changed, clear, item_id = itemId.Length > 0 ? itemId : null };
            });
        }

        /// <summary>
        /// レイヤー単位でアイテムの重なりを解消する。
        /// 各レイヤーをframe昇順に並べ、後続アイテムが前のアイテムのend(frame+length)に重なる場合、後ろへシフトする。
        /// gap で各アイテム間に最小すき間（フレーム）を確保できる。
        /// </summary>
        private async Task<object> ResolveOverlaps(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int gap = TimelineInputValidation.AtLeast(GetInt(b, "gap", 0), 0, "gap");
            int[]? onlyLayers = TimelineInputValidation.GetLayers(b);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                var candidates = new List<(object item, int frame, int length, int layer)>();
                foreach (var iv in rawItems)
                {
                    var (frame, layer, length, _, _, item) = ReadItemInfo(iv);
                    if (onlyLayers != null && !onlyLayers.Contains(layer)) continue;
                    candidates.Add((item, frame, length, layer));
                }

                var moved = new List<object>();
                TimelineInputValidation.ApplyResolve(candidates, gap, (item, frame, newFrame, length, layer) =>
                {
                    try
                    {
                        var fProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        fProp?.SetValue(item, newFrame);
                        moved.Add(new { layer, oldFrame = frame, newFrame, length });
                    }
                    catch (Exception ex) { moved.Add(new { layer, oldFrame = frame, error = ex.Message }); }
                });
                return (object)new { success = true, movedCount = moved.Count, moved };
            });
        }

        /// <summary>
        /// 指定フレーム以降のアイテムを一括で前後にシフトする（挿入・詰めに使う）。
        /// fromFrame 以上のアイテムの Frame に delta を加算する。layers 指定でレイヤーを絞れる。
        /// </summary>
        private async Task<object> ShiftItems(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int fromFrame = TimelineInputValidation.AtLeast(GetInt(b, "fromFrame", 0), 0, "fromFrame");
            int delta = GetInt(b, "delta", 0);
            int[]? onlyLayers = TimelineInputValidation.GetLayers(b);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                var candidates = new List<(object item, int frame, int length, int layer)>();
                foreach (var iv in rawItems)
                {
                    var (frame, layer, length, _, _, item) = ReadItemInfo(iv);
                    if (frame < fromFrame) continue;
                    if (onlyLayers != null && !onlyLayers.Contains(layer)) continue;
                    candidates.Add((item, frame, length, layer));
                }
                var moved = new List<object>();
                TimelineInputValidation.ApplyShift(candidates, delta, (item, frame, newFrame, layer) =>
                {
                    try
                    {
                        var fProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        fProp?.SetValue(item, newFrame);
                        moved.Add(new { layer, oldFrame = frame, newFrame });
                    }
                    catch (Exception ex) { moved.Add(new { layer, oldFrame = frame, error = ex.Message }); }
                });
                return (object)new { success = true, movedCount = moved.Count, fromFrame, delta, moved };
            });
        }

        /// <summary>指定アイテム(frame+layer)に適用されている全エフェクトとそのパラメータ現在値を取得する。</summary>
        private object GetItemEffects(HttpListenerRequest req)
        {
            int tf = -1, tl = -1;
            int.TryParse(req.QueryString["frame"], out tf);
            int.TryParse(req.QueryString["layer"], out tl);
            string itemId = req.QueryString["item_id"] ?? "";
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                object? targetItem = null;
                if (itemId.Length > 0)
                {
                    targetItem = FindItemById(rawItems, itemId, out bool ambiguous);
                    if (ambiguous) return Failure("ITEM_ID_AMBIGUOUS", "item_id が複数のアイテムに一致しました");
                }
                else
                {
                    foreach (var iv in rawItems)
                    {
                        var (frame, layer, _, _, _, item) = ReadItemInfo(iv);
                        if ((tf < 0 || frame == tf) && (tl < 0 || layer == tl)) { targetItem = item; break; }
                    }
                }
                if (targetItem == null) return Failure("ITEM_NOT_FOUND", itemId.Length > 0 ? "item_id に一致するアイテムがありません" : $"アイテム未発見 frame={tf} layer={tl}");

                var result = new Dictionary<string, object>();
                foreach (var collName in new[] { "VideoEffects", "AudioEffects", "Effects" })
                {
                    var coll = GetPropObj(targetItem, collName) as System.Collections.IEnumerable;
                    if (coll == null) continue;
                    var effects = new List<object>();
                    foreach (var eff in coll)
                    {
                        if (eff == null) continue;
                        var et = eff.GetType();
                        var pvals = new List<object>();
                        foreach (var p in et.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                        {
                            if (p.GetIndexParameters().Length > 0) continue;
                            object? v = null; try { v = p.GetValue(eff); } catch { }
                            pvals.Add(new { name = p.Name, type = p.PropertyType.Name, value = ToJsonSafe(v) });
                        }
                        effects.Add(new { type = et.Name, fullType = et.FullName, properties = pvals });
                    }
                    if (effects.Count > 0) result[collName] = effects;
                }
                var identity = GetItemIdentity(targetItem);
                var info = ReadItemInfo(targetItem);
                return (object)new
                {
                    success = true,
                    item_id = identity.id,
                    revision = GetItemRevision(targetItem),
                    identity_persistent = identity.persistent,
                    frame = info.frame,
                    layer = info.layer,
                    itemType = targetItem.GetType().Name,
                    effects = result
                };
            });
        }

        private static void Start_SleepMs(int ms) => System.Threading.Thread.Sleep(ms);

        private object? GetMainModel(object vm)
        {
            foreach (var f in vm.GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance))
                if (f.FieldType.Name.Contains("MainModel")) { var v = f.GetValue(vm); if (v != null) return v; }
            foreach (var p in vm.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
                if (p.PropertyType.Name.Contains("MainModel")) { try { var v = p.GetValue(vm); if (v != null) return v; } catch { } }
            return null;
        }

        // アイテムにVideoEffectを追加
        private async Task<object> AddVideoEffect(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            TimelineInputValidation.RequireCoordinates(b);
            int tf = GetInt(b, "frame", 0); int tl = GetInt(b, "layer", 0);
            string eName = GetStr(b, "effect", "");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };
                object? targetItem = null;
                foreach (var iv in rawItems) { var item = GetPropObj(iv, "Item") ?? iv; try { int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); if (f2 == tf && l2 == tl) { targetItem = item; break; } } catch { } }
                if (targetItem == null) return (object)new { success = false, error = $"アイテム未発見 f={tf} l={tl}" };
                var eType = AppDomain.CurrentDomain.GetAssemblies().SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } }).FirstOrDefault(t => !t.IsAbstract && !t.IsInterface && (t.FullName == eName || t.Name == eName || t.Name.Contains(eName)));
                if (eType == null) return (object)new { success = false, error = $"エフェクト型未発見: {eName}" };
                var veProp = targetItem.GetType().GetProperty("VideoEffects", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (veProp == null) return (object)new { success = false, error = "VideoEffectsなし" };
                try
                {
                    var curList = veProp.GetValue(targetItem);
                    var eObj = Activator.CreateInstance(eType);
                    // パラメータ設定
                    if (b.TryGetValue("params", out var pe) && eObj != null)
                        foreach (var kv in pe.EnumerateObject()) { var p = eType.GetProperty(kv.Name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); if (p == null) continue; try { object? v = kv.Value.ValueKind == System.Text.Json.JsonValueKind.Number ? Convert.ChangeType(kv.Value.GetDouble(), p.PropertyType) : (object?)kv.Value.GetString(); p.SetValue(eObj, v); } catch { } }
                    var addM = curList?.GetType().GetMethod("Add", BindingFlags.Public | BindingFlags.Instance);
                    if (addM == null) return (object)new { success = false, error = "ImmutableList.Addなし" };
                    veProp.SetValue(targetItem, addM.Invoke(curList, new[] { eObj }));
                    MarkItemChanged(targetItem);
                    return (object)new { success = true, effect = eType.FullName, revision = GetItemRevision(targetItem) };
                }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            });
        }

        // アイテムの単一プロパティを設定（FadeIn/FadeOutなど）
        private async Task<object> SetItemProp(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int tf = GetInt(b, "frame", 0); int tl = GetInt(b, "layer", 0);
            string itemId = GetStr(b, "item_id", "");
            string expectedRevision = GetStr(b, "expected_revision", "");
            if (expectedRevision.Length > 0 && itemId.Length == 0)
                return Failure("REVISION_REQUIRES_ITEM_ID", "expected_revision を使う場合は item_id も指定してください");
            TimelineInputValidation.RequireItemTarget(b);
            string pName = GetStr(b, "prop", ""); string pVal = GetStr(b, "value", "");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return Failure("NO_MAIN_VIEW_MODEL", "VM失敗");
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return Failure("NO_TIMELINE", "TVM失敗");
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return Failure("ITEMS_UNAVAILABLE", "Items失敗");
                object? targetItem = null;
                if (itemId.Length > 0)
                {
                    targetItem = FindItemById(rawItems, itemId, out bool ambiguous);
                    if (ambiguous) return Failure("ITEM_ID_AMBIGUOUS", "item_id が複数のアイテムに一致しました");
                }
                else
                {
                    foreach (var iv in rawItems)
                    {
                        var (frame, layer, _, _, _, item) = ReadItemInfo(iv);
                        if (frame == tf && layer == tl) { targetItem = item; break; }
                    }
                }
                if (targetItem == null) return Failure("ITEM_NOT_FOUND", itemId.Length > 0 ? "item_id に一致するアイテムがありません" : $"アイテム未発見 f={tf} l={tl}");
                var conflict = RevisionConflict(targetItem, expectedRevision);
                if (conflict != null) return conflict;
                var p = targetItem.GetType().GetProperty(pName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (p == null || !p.CanWrite) return Failure("PROPERTY_NOT_WRITABLE", $"プロパティ'{pName}'を変更できません");
                object? converted;
                try { converted = Convert.ChangeType(pVal, p.PropertyType); }
                catch (Exception ex) { return Failure("INVALID_ARGUMENT", ex.InnerException?.Message ?? ex.Message); }
                try
                {
                    p.SetValue(targetItem, converted);
                }
                catch (Exception ex)
                {
                    MarkItemChanged(targetItem); // setter may have changed state before throwing
                    return Failure("PROPERTY_SET_FAILED", ex.InnerException?.Message ?? ex.Message, true);
                }
                MarkItemChanged(targetItem);
                object? actual;
                try { actual = p.GetValue(targetItem); }
                catch (Exception ex) { return Failure("PROPERTY_VERIFY_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
                var identity = GetItemIdentity(targetItem);
                if (!Equals(converted, actual))
                    return (object)new { success = false, error_code = "PROPERTY_VERIFY_FAILED",
                        error = $"プロパティ'{pName}'の設定後の値が要求と一致しません", retryable = false,
                        outcome_unknown = true, item_id = identity.id, revision = GetItemRevision(targetItem),
                        expected = ToJsonSafe(converted), actual = ToJsonSafe(actual) };
                return (object)new { success = true, item_id = identity.id, revision = GetItemRevision(targetItem),
                    identity_persistent = identity.persistent, prop = pName, value = pVal, verified = true,
                    actual = ToJsonSafe(actual) };
            });
        }

        // 使えるエフェクト一覧
        private object ListEffects()
        {
            var effects = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                .Where(t => !t.IsAbstract && !t.IsInterface && t.Namespace?.StartsWith("YukkuriMovieMaker.Project.Effects") == true)
                .Select(t => new { name = t.Name.Replace("Effect", ""), fullName = t.Name })
                .OrderBy(t => t.name).ToArray();
            return new { count = effects.Length, effects };
        }

        private object DescribeEffect(HttpListenerRequest req)
        {
            string name = req.QueryString["name"] ?? "";
            if (string.IsNullOrWhiteSpace(name)) return Failure("INVALID_ARGUMENT", "name パラメータが必要です");
            var matches = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                .Where(t => !t.IsAbstract && !t.IsInterface && t.Namespace?.StartsWith("YukkuriMovieMaker.Project.Effects") == true)
                .Where(t => t.Name == name || t.FullName == name || t.Name == name + "Effect")
                .Take(2).ToArray();
            if (matches.Length == 0) return Failure("EFFECT_NOT_FOUND", $"エフェクト '{name}' が見つかりません");
            if (matches.Length > 1) return Failure("EFFECT_AMBIGUOUS", "同名のエフェクトがあります。fullName を指定してください");
            return EffectMetadata.Describe(matches[0]);
        }

        // アイテム削除（item_id / item_ids / layer指定 / frame+layer指定）
        private async Task<object> DeleteItems(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int[]? layers = TimelineInputValidation.GetLayers(b);
            int targetFrame = TimelineInputValidation.AtLeast(GetInt(b, "frame", -1), -1, "frame");
            int targetLayer = TimelineInputValidation.AtLeast(GetInt(b, "layer", -1), -1, "layer");
            string itemId = GetStr(b, "item_id", "");
            string expectedRevision = GetStr(b, "expected_revision", "");
            var requestedIds = new List<string>();
            if (itemId.Length > 0) requestedIds.Add(itemId);
            if (b.TryGetValue("item_ids", out var idsEl) && idsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var el in idsEl.EnumerateArray())
                {
                    if (el.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("item_ids must be strings");
                    string id = el.GetString() ?? "";
                    if (id.Length == 0) throw new ArgumentException("item_ids must not contain empty ids");
                    if (!requestedIds.Contains(id)) requestedIds.Add(id);
                }
            }
            if (expectedRevision.Length > 0 && requestedIds.Count != 1)
                return Failure("REVISION_REQUIRES_ITEM_ID", "expected_revision を使う場合は item_id を1件指定してください");

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };

                var toRemove = new List<object>();
                var missing = new List<string>();
                if (requestedIds.Count > 0)
                {
                    foreach (string id in requestedIds)
                    {
                        var target = FindItemById(rawItems, id, out bool ambiguous);
                        if (ambiguous) return Failure("ITEM_ID_AMBIGUOUS", "item_id が複数のアイテムに一致しました: " + id);
                        if (target == null)
                        {
                            missing.Add(id);
                            continue;
                        }
                        var conflict = RevisionConflict(target, expectedRevision);
                        if (conflict != null) return conflict;
                        toRemove.Add(target);
                    }
                    if (toRemove.Count == 0 && requestedIds.Count == 1)
                        return Failure("ITEM_NOT_FOUND", "item_id に一致するアイテムがありません");
                }
                else
                {
                    foreach (var iv in rawItems)
                    {
                        var (frame, layer, _, _, _, item) = ReadItemInfo(iv);
                        if (layers != null && layers.Contains(layer)) toRemove.Add(item);
                        else if (targetFrame >= 0 && targetLayer >= 0 && frame == targetFrame && layer == targetLayer) toRemove.Add(item);
                    }
                }
                if (toRemove.Count == 0)
                    return (object)new { success = true, removed = 0, item_ids = Array.Empty<string>(), missing = missing.ToArray(), note = "対象アイテムなし" };
                var deleted = TryRemoveTimelineItems(tvm, toRemove, missing);
                return deleted.payload;
            });
        }

        private (bool ok, object payload) TryRemoveTimelineItems(object tvm, List<object> toRemove, IReadOnlyList<string>? missing = null)
        {
            var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
            var timelineObj = timelineField?.GetValue(tvm);
            if (timelineObj == null) return (false, (object)new { success = false, error = "timelineフィールド取得失敗" });

            var deleteMethod = timelineObj.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .FirstOrDefault(m => m.Name == "DeleteItems" && m.GetParameters().Length == 1);
            if (deleteMethod == null) return (false, (object)new { success = false, error = "DeleteItems未発見" });
            try
            {
                var iItemType = toRemove[0].GetType().GetInterfaces().FirstOrDefault(i => i.Name == "IItem");
                Array arr;
                if (iItemType != null)
                {
                    arr = Array.CreateInstance(iItemType, toRemove.Count);
                    for (int i = 0; i < toRemove.Count; i++) arr.SetValue(toRemove[i], i);
                }
                else { arr = toRemove.ToArray(); }
                var removedIds = toRemove.Select(item => GetItemIdentity(item).id).ToArray();
                deleteMethod.Invoke(timelineObj, new object[] { arr });
                ForgetEditBindings(removedIds);
                return (true, (object)new { success = true, removed = toRemove.Count, item_ids = removedIds,
                    missing = missing?.ToArray() ?? Array.Empty<string>() });
            }
            catch (Exception ex)
            {
                return (false, (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message });
            }
        }

        private async Task<object> AddEffectToItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            TimelineInputValidation.RequireCoordinates(b);
            // target: "voice"/"image"/"tachie", targetFrame, targetLayer, effectName, params...
            string effectName = GetStr(b, "effect", "");
            int targetFrame = GetInt(b, "frame", 0);
            int targetLayer = GetInt(b, "layer", 0);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                // 対象アイテムを探す
                object? targetItem = null;
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var fProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    var lProp = item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    try
                    {
                        int f2 = (int)(fProp?.GetValue(item) ?? -1);
                        int l2 = (int)(lProp?.GetValue(item) ?? -1);
                        if (f2 == targetFrame && l2 == targetLayer) { targetItem = item; break; }
                    }
                    catch { }
                }
                if (targetItem == null) return (object)new { success = false, error = $"frame={targetFrame} layer={targetLayer} のアイテムが見つかりません" };

                // エフェクトコレクションを取得してエフェクト追加
                var effectsProp = targetItem.GetType().GetProperty("Effects", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (effectsProp == null) effectsProp = targetItem.GetType().GetProperty("VideoEffects", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (effectsProp == null) effectsProp = targetItem.GetType().GetProperty("AudioEffects", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (effectsProp == null)
                {
                    var props = targetItem.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Select(p => p.Name + ":" + p.PropertyType.Name).ToArray();
                    return (object)new { success = false, error = "Effectsプロパティ見つからず", props };
                }

                var effects = effectsProp.GetValue(targetItem);
                var addMethod = effects?.GetType().GetMethod("Add", BindingFlags.Public | BindingFlags.Instance);
                if (addMethod == null) return (object)new { success = false, error = "effects.Add見つからず", effectsType = effects?.GetType().Name };

                // エフェクト型を名前で検索
                var effectType = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                    .FirstOrDefault(t => (t.Name.Contains(effectName) || (t.GetCustomAttributes(false).Any(a => a.ToString()?.Contains(effectName) == true))) && !t.IsAbstract && !t.IsInterface);

                if (effectType == null) return (object)new { success = false, error = $"エフェクト型 '{effectName}' が見つかりません" };

                try
                {
                    var effectObj = Activator.CreateInstance(effectType);
                    var invokeResult = addMethod.Invoke(effects, new[] { effectObj });
                    if (addMethod.ReturnType != typeof(void) && invokeResult != null)
                    {
                        effectsProp.SetValue(targetItem, invokeResult);
                    }
                    MarkItemChanged(targetItem);
                    return (object)new { success = true, effect = effectType.Name, item = targetItem.GetType().Name, effectsType = effects?.GetType().FullName, revision = GetItemRevision(targetItem) };
                }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            });
        }

        private object DebugTachie()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var mainModel = GetMainModel(vm);
                if (mainModel == null) return (object)new { error = "MainModel取得失敗" };
                var methods = mainModel.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => m.Name.Contains("Tachie") || m.Name.Contains("Face") || m.Name.Contains("Effect") || m.Name.Contains("Audio"))
                    .Select(m => m.Name + "(" + string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name)) + ")")
                    .OrderBy(s => s).ToArray();
                return (object)new { methods };
            });
        }

        private object DebugFaceItemProps()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { error = "Items失敗" };
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    if (!item.GetType().Name.Contains("Face")) continue;
                    var props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Select(p => { string? v = null; try { var val = p.GetValue(item); v = val?.ToString(); if (val != null && v != null && v.StartsWith(val.GetType().Namespace ?? "")) { var subProps = val.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(sp => { try { return sp.Name + "=" + sp.GetValue(val); } catch { return sp.Name + "=?"; } }); v = "{" + string.Join(", ", subProps) + "}"; } } catch { } return new { p.Name, type = p.PropertyType.Name, v }; })
                        .ToArray();
                    return (object)new { typeName = item.GetType().FullName, props };
                }
                return (object)new { error = "TachieFaceItemなし" };
            });
        }

        private async Task<object> DeleteItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int tf = GetInt(b, "frame", -1); int tl = GetInt(b, "layer", -1);
            string typePat = GetStr(b, "type", "");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var mm = GetMainModel(vm);
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };
                var toDelete = new System.Collections.Generic.List<object>();
                foreach (var iv in rawItems) { var item = GetPropObj(iv, "Item") ?? iv; try { int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); bool match = (tf == -1 || f2 == tf) && (tl == -1 || l2 == tl) && (typePat == "" || item.GetType().Name.Contains(typePat)); if (match) toDelete.Add(iv); } catch { } }
                if (toDelete.Count == 0) return (object)new { success = false, error = "対象アイテムなし" };
                int deleted = 0;
                foreach (var iv in toDelete)
                {
                    try
                    {
                        if (mm != null) { var item = GetPropObj(iv, "Item") ?? iv; var rm = mm.GetType().GetMethod("RemoveItem", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); if (rm != null) { rm.Invoke(mm, new[] { item }); deleted++; continue; } }
                        var tl2 = GetPropObj(vm, "Timeline") ?? GetPropObj(tvm, "Timeline"); if (tl2 != null) { var tryRm = tl2.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).FirstOrDefault(m => m.Name.Contains("Remove")); if (tryRm != null) { tryRm.Invoke(tl2, new[] { GetPropObj(iv, "Item") ?? iv }); deleted++; } }
                    }
                    catch { }
                }
                return (object)new { success = deleted > 0, deleted, total = toDelete.Count };
            });
        }

        // 表情パラメータを変更（FacePath等）
        private async Task<object> SetFaceParam(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            TimelineInputValidation.RequireItemTarget(b);
            int tf = GetInt(b, "frame", -1); int tl = GetInt(b, "layer", -1);
            string itemId = GetStr(b, "item_id", "");
            string expectedRevision = GetStr(b, "expected_revision", "");
            if (expectedRevision.Length > 0 && itemId.Length == 0)
                return Failure("REVISION_REQUIRES_ITEM_ID", "expected_revision を使う場合は item_id も指定してください");
            if (!b.Keys.Any(k => k is not ("frame" or "layer" or "item_id" or "expected_revision")))
                throw new ArgumentException("at least one face parameter is required");
            // keyValuePairs: { "FacePath": "...", etc. }
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var located = LocateTimelineItem(itemId, tf, tl);
                if (located.error != null) return located.error;
                var targetItem = located.item!;
                if (!targetItem.GetType().Name.Contains("Face", StringComparison.OrdinalIgnoreCase))
                    return Failure("FACE_ITEM_REQUIRED", "対象はFaceItemではありません");
                var conflict = RevisionConflict(targetItem, expectedRevision);
                if (conflict != null) return conflict;
                // FaceParameterを探して設定
                var fpProp = targetItem.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(p => p.Name.Contains("FaceParameter") || p.Name.Contains("FaceParam"));
                var faceParam = fpProp?.GetValue(targetItem);
                var results = new System.Collections.Generic.List<string>();
                var pending = new System.Collections.Generic.List<(PropertyInfo property, object owner, object? value, string name)>();
                // Resolve and convert every requested parameter before setting any of them.
                // A typo or invalid value must not leave a partially edited face item.
                foreach (var kv in b)
                {
                    if (kv.Key is "frame" or "layer" or "item_id" or "expected_revision") continue;
                    var directProp = targetItem.GetType().GetProperty(kv.Key, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    var nestedProp = faceParam?.GetType().GetProperty(kv.Key, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    var prop = directProp ?? nestedProp;
                    var owner = directProp != null ? targetItem : faceParam;
                    if (prop == null || owner == null || !prop.CanWrite || prop.GetIndexParameters().Length != 0)
                    {
                        results.Add($"{kv.Key}=NOTFOUND");
                        continue;
                    }
                    try
                    {
                        var raw = kv.Value.ValueKind == JsonValueKind.String ? kv.Value.GetString() : kv.Value.ToString();
                        var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                        object? value = type.IsEnum ? Enum.Parse(type, raw ?? "", true) : Convert.ChangeType(raw, type);
                        pending.Add((prop, owner, value, kv.Key));
                    }
                    catch (Exception ex) { results.Add($"{kv.Key}=NG({ex.Message})"); }
                }
                if (results.Count > 0)
                    return (object)new { success = false, changed = false, error_code = "FACE_PARAM_INVALID", results };
                foreach (var change in pending)
                {
                    try { change.property.SetValue(change.owner, change.value); results.Add($"{change.name}=OK"); }
                    catch (Exception ex) { results.Add($"{change.name}=NG({ex.Message})"); break; }
                }
                bool changed = results.Any(result => result.EndsWith("=OK", StringComparison.Ordinal));
                if (changed)
                {
                    MarkItemChanged(targetItem);
                    var vm = GetMainViewModel();
                    var model = vm == null ? null : GetMainModel(vm);
                    if (model != null) TryRecordHistory(model);
                }
                bool complete = changed && results.Count == pending.Count &&
                    results.All(result => result.EndsWith("=OK", StringComparison.Ordinal));
                return (object)new { success = complete, changed, item_id = GetItemIdentity(targetItem).id,
                    revision = GetItemRevision(targetItem), results,
                    error_code = complete ? null : changed ? "FACE_PARAM_PARTIAL" : "FACE_PARAM_UNCHANGED",
                    outcome_unknown = changed && !complete };
            });
        }

        // VoiceItemのプロパティ（AudioEffects等）をデバッグ
        private object DebugVoiceItemProps()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { error = "Items失敗" };

                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    if (!item.GetType().Name.Contains("Voice")) continue;
                    var props = item.GetType()
                        .GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Select(p =>
                        {
                            string? val = null;
                            string? innerType = null;
                            try
                            {
                                var v = p.GetValue(item);
                                val = v?.ToString();
                                if (v != null && p.PropertyType.IsGenericType)
                                    innerType = string.Join(", ", p.PropertyType.GetGenericArguments().Select(t => t.FullName));
                            }
                            catch { }
                            return new { p.Name, TypeName = p.PropertyType.Name, innerType, val };
                        }).ToArray();
                    return (object)new { typeName = item.GetType().FullName, props };
                }
                return (object)new { error = "VoiceItemが見つかりません" };
            });
        }

        // 音声エフェクト追加
        private async Task<object> AddAudioEffect(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int frame = GetInt(b, "frame", -1);
            int layer = GetInt(b, "layer", -1);
            string effect = GetStr(b, "effect", "");

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };

                // 対象アイテムを検索
                object? targetItem = null;
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    try
                    {
                        int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1);
                        int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1);
                        if (f2 == frame && l2 == layer) { targetItem = item; break; }
                    }
                    catch { }
                }
                if (targetItem == null) return (object)new { success = false, error = $"frame={frame} layer={layer} のアイテムが見つかりません" };

                // AudioEffectsプロパティを取得
                var audioEffProp = targetItem.GetType().GetProperty("AudioEffects",
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (audioEffProp == null)
                    return (object)new { success = false, error = "AudioEffectsプロパティが見つかりません", type = targetItem.GetType().Name };

                var audioEffects = audioEffProp.GetValue(targetItem);
                if (audioEffects == null) return (object)new { success = false, error = "AudioEffectsがnull" };

                // エフェクト型を検索
                var allTypes = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return Array.Empty<Type>(); } });
                var effType = allTypes.FirstOrDefault(t =>
                    !t.IsAbstract && !t.IsInterface &&
                    (t.Name.Equals(effect, StringComparison.OrdinalIgnoreCase) ||
                     t.Name.Equals(effect + "Effect", StringComparison.OrdinalIgnoreCase) ||
                     t.Name.Contains(effect, StringComparison.OrdinalIgnoreCase)) &&
                    t.GetInterfaces().Any(i => i.Name.Contains("AudioEffect") || i.Name.Contains("IAudio")));

                if (effType == null)
                    return (object)new { success = false, error = $"音声エフェクト型'{effect}'が見つかりません" };

                var effInstance = Activator.CreateInstance(effType);
                if (effInstance == null) return (object)new { success = false, error = "エフェクトインスタンス生成失敗" };

                // ImmutableList.Addパターン
                var addMethod = audioEffects.GetType().GetMethod("Add");
                if (addMethod != null)
                {
                    var newList = addMethod.Invoke(audioEffects, new[] { effInstance });
                    audioEffProp.SetValue(targetItem, newList);
                    MarkItemChanged(targetItem);
                }
                else return (object)new { success = false, error = "Addメソッドが見つかりません", listType = audioEffects.GetType().FullName };

                return (object)new { success = true, effect = effType.FullName, revision = GetItemRevision(targetItem) };
            });
        }

        private object DebugTachieItemProps()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { error = "Items取得失敗" };

                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    if (!item.GetType().Name.Contains("Tachie")) continue;
                    var props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Select(p => { string? val = null; try { val = p.GetValue(item)?.ToString(); } catch { } return new { p.Name, TypeName = p.PropertyType.Name, val }; })
                        .ToArray();
                    return (object)new { typeName = item.GetType().FullName, props };
                }
                return (object)new { error = "TachieItemが見つかりません" };
            });
        }

        private object DebugFaceTypes()
        {
            // 表情・登場退場エフェクト関連の型を列挙
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                .Where(t => !t.IsAbstract && !t.IsInterface && (
                    t.Name.Contains("Face") || t.Name.Contains("Appear") || t.Name.Contains("Disappear") ||
                    t.Name.Contains("Enter") || t.Name.Contains("Exit") || t.Name.Contains("登場") ||
                    t.Name.Contains("退場") || t.Name.Contains("Motion") || t.Name.Contains("Tachie") ||
                    (t.GetInterfaces().Any(i => i.Name == "IVideoEffect") && !t.IsNested)))
                .Select(t => new { t.FullName, interfaces = string.Join(",", t.GetInterfaces().Select(i => i.Name)) })
                .OrderBy(t => t.FullName)
                .ToArray();
            return new { count = types.Length, types };
        }

        private async Task<object> ArrangeItems(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            if (!body.TryGetValue("order", out var orderElem))
                return new { success = false, error = "order パラメータが必要です" };
            var order = orderElem.EnumerateArray().Select(e => e.GetString() ?? "").ToList();
            int targetLayer = GetInt(body, "layer", 0);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                var itemMap = new Dictionary<string, object>();
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var fp = GetPropObj(item, "FilePath")?.ToString() ?? "";
                    var fn = System.IO.Path.GetFileName(fp);
                    if (!string.IsNullOrEmpty(fn)) itemMap[fn] = item;
                }

                var results = new List<object>();
                int currentFrame = 0;
                for (int i = 0; i < order.Count; i++)
                {
                    var name = order[i];
                    var key = itemMap.Keys.FirstOrDefault(k => k.Contains(name) || name.Contains(k));
                    if (key == null) { results.Add(new { name, success = false, error = "見つかりません" }); continue; }
                    var item = itemMap[key];
                    try
                    {
                        int length = 300;
                        try { length = (int)(GetPropObj(item, "Length") ?? 300); } catch { }
                        item.GetType().GetProperty("Layer")?.SetValue(item, targetLayer);
                        item.GetType().GetProperty("Frame")?.SetValue(item, currentFrame);
                        results.Add(new { name = key, success = true, layer = targetLayer, frame = currentFrame, length });
                        currentFrame += length;
                    }
                    catch (Exception ex) { results.Add(new { name = key, success = false, error = ex.Message }); }
                }
                return (object)new { success = true, results };
            });
        }

        private async Task<object> ReorderItems(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            if (!body.TryGetValue("order", out var orderElem))
                return new { success = false, error = "order パラメータが必要です" };
            var order = orderElem.EnumerateArray().Select(e => e.GetString() ?? "").ToList();

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { success = false, error = "Items取得失敗" };

                var itemMap = new Dictionary<string, object>();
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var fp = GetPropObj(item, "FilePath")?.ToString() ?? "";
                    var fn = System.IO.Path.GetFileName(fp);
                    if (!string.IsNullOrEmpty(fn)) itemMap[fn] = item;
                }

                var results = new List<object>();
                for (int i = 0; i < order.Count; i++)
                {
                    var name = order[i];
                    var key = itemMap.Keys.FirstOrDefault(k => k.Contains(name) || name.Contains(k));
                    if (key == null) { results.Add(new { name, success = false, error = "見つかりません" }); continue; }
                    var item = itemMap[key];
                    try { item.GetType().GetProperty("Layer")?.SetValue(item, i); results.Add(new { name = key, success = true, newLayer = i }); }
                    catch (Exception ex) { results.Add(new { name = key, success = false, error = ex.Message }); }
                }
                return (object)new { success = true, results };
            });
        }

        private object ExecCommand(string name)
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                return TryCmd(vm, name) ? (object)new { success = true } : new { success = false, error = $"'{name}' が見つかりません" };
            });
        }

        // ════════════════════════════════════════════════════════════
        // ■ 全機能アクセス用 汎用API
        //   YMM4内部の任意のコマンド・プロパティ・メソッドにアクセスする。
        //   個別ハンドラで未対応の機能でも、これらの汎用APIで操作できる。
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 対象オブジェクト(Main/ActiveTimeline/Player/Project/Selection)を解決する共通ヘルパー。
        /// path で "ActiveTimeline.Items[0].Item" のようなドット/インデックス指定の深掘りも可能。
        /// </summary>
        private static object? ResolveTarget(string target)
        {
            var vm = GetMainViewModel();
            if (vm == null) return null;

            // ベースオブジェクトの解決
            string baseName = target;
            string rest = "";
            int dot = target.IndexOf('.');
            if (dot >= 0) { baseName = target.Substring(0, dot); rest = target.Substring(dot + 1); }

            object? obj = baseName switch
            {
                "Main" or "MainViewModel" or "" => vm,
                "ActiveTimeline" or "ActiveTimelineViewModel" or "Timeline" => GetPropObj(vm, "ActiveTimelineViewModel"),
                "Player" or "PlayerViewModel" => GetPreviewViewModel()
                    ?? GetPropObj(vm, "PlayerViewModel") ?? GetPropObj(vm, "Player") ?? GetPropObj(vm, "player") ?? GetPropObj(vm, "_player"),
                "Preview" or "PreviewViewModel" => GetPreviewViewModel(),
                "Project" => GetPropObj(vm, "Project") ?? GetPropObj(vm, "project") ?? GetPropObj(vm, "_project"),
                _ => GetPropObj(vm, baseName)
            };

            if (obj == null || string.IsNullOrEmpty(rest)) return obj;
            return ResolvePath(obj, rest);
        }

        /// <summary>"A.B[2].C" のようなパスを辿って値を取得する（ReactivePropertyの.Valueも自動展開）。</summary>
        private static object? ResolvePath(object? obj, string path)
        {
            return ReflectionPath.TryResolve(obj, path, out var value) ? value : null;
        }

        /// <summary>任意の ICommand をターゲット上で実行する。UI からしか押せないコマンドを直接トリガーできる。</summary>
        private async Task<object> ExecCommandApi(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string name = GetStr(b, "name", "");
            string target = GetStr(b, "target", "Main");
            // param は文字列/数値/null のいずれか
            object? param = null;
            if (b.TryGetValue("param", out var pe))
            {
                param = pe.ValueKind switch
                {
                    JsonValueKind.String => pe.GetString(),
                    JsonValueKind.Number => pe.TryGetInt32(out var iv) ? iv : pe.GetDouble(),
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    _ => null
                };
            }
            if (string.IsNullOrEmpty(name)) return Failure("INVALID_ARGUMENT", "name パラメータが必要です");

            if (!_allowAdvanced && !((target == "Main" && (name == "UndoCommand" || name == "RedoCommand")) ||
                (target == "ActiveTimeline" && (name == "SplitItemCommand" || name == "AlignItemsCommand"))))
                return Failure("ADVANCED_DISABLED", "This command requires advanced APIs");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var obj = ResolveTarget(target);
                if (obj == null) return Failure("TARGET_NOT_FOUND", $"target '{target}' 取得失敗");

                // プロパティ・フィールド両方から ICommand を探す
                var t = obj.GetType();
                var cmdProp = t.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                object? cmdObj = null;
                try { cmdObj = cmdProp?.GetValue(obj); } catch { }
                if (cmdObj == null)
                {
                    var cmdField = t.GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    try { cmdObj = cmdField?.GetValue(obj); } catch { }
                }
                if (cmdObj is not ICommand cmd)
                    return Failure("COMMAND_NOT_FOUND", $"'{name}' は ICommand ではありません（target={target}）");

                try
                {
                    bool canExec = cmd.CanExecute(param);
                    if (!canExec) return (object)new { success = false, error_code = "COMMAND_UNAVAILABLE", error = $"'{name}'.CanExecute=false（現在実行不可）", retryable = false, outcome_unknown = false, canExecute = false };
                    cmd.Execute(param);
                    return (object)new { success = true, command = name, target };
                }
                catch (Exception ex) { return Failure("COMMAND_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
            });
        }

        /// <summary>ターゲット上の ICommand 一覧を返す（どんなコマンドが使えるか発見するため）。</summary>
        private object ListCommands(string target = "")
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var result = new Dictionary<string, object>();
                foreach (var tgt in new[] { "Main", "ActiveTimeline", "Player", "Project" })
                {
                    var obj = ResolveTarget(tgt);
                    if (obj == null) continue;
                    var t = obj.GetType();
                    var cmds = new List<(string name, bool can)>();
                    foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
                    {
                        if (!typeof(ICommand).IsAssignableFrom(p.PropertyType)) continue;
                        bool can = false;
                        try { if (p.GetValue(obj) is ICommand c) can = c.CanExecute(null); } catch { }
                        cmds.Add((p.Name, can));
                    }
                    result[tgt] = new
                    {
                        type = t.Name,
                        commands = cmds.OrderBy(c => c.name).Select(c => new { name = c.name, canExecuteNow = c.can }).ToList()
                    };
                }
                return (object)result;
            });
        }

        /// <summary>任意のオブジェクトの任意プロパティ/フィールドを取得（ReactivePropertyは.Value展開）。</summary>
        private async Task<object> ReflectGet(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string target = GetStr(b, "target", "Main");
            string path = GetStr(b, "path", "");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var baseObj = ResolveTarget(target);
                if (baseObj == null) return Failure("TARGET_NOT_FOUND", $"target '{target}' 取得失敗");
                if (!ReflectionPath.TryResolve(baseObj, path, out var val))
                    return Failure("PATH_NOT_FOUND", $"path '{path}' 取得失敗");
                return (object)new { success = true, target, path, type = val?.GetType().FullName, value = ToJsonSafe(val) };
            });
        }

        /// <summary>任意のオブジェクトの任意プロパティ/フィールドを設定（ReactivePropertyの.Valueも対応）。</summary>
        private async Task<object> ReflectSet(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string target = GetStr(b, "target", "Main");
            string path = GetStr(b, "path", "");
            if (string.IsNullOrEmpty(path)) return Failure("INVALID_ARGUMENT", "path パラメータが必要です");
            if (!b.TryGetValue("value", out var valElem)) return Failure("INVALID_ARGUMENT", "value パラメータが必要です");

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var baseObj = ResolveTarget(target);
                if (baseObj == null) return Failure("TARGET_NOT_FOUND", $"target '{target}' 取得失敗");

                // 最後のセグメントの親オブジェクトを取得
                object? parent = baseObj;
                string lastSeg = path;
                int lastDot = path.LastIndexOf('.');
                if (lastDot >= 0)
                {
                    parent = ResolvePath(baseObj, path.Substring(0, lastDot));
                    lastSeg = path.Substring(lastDot + 1);
                }
                if (parent == null) return Failure("PATH_NOT_FOUND", "親オブジェクト取得失敗");

                var prop = parent.GetType().GetProperty(lastSeg, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                var field = prop == null ? parent.GetType().GetField(lastSeg, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance) : null;
                if (prop == null && field == null) return Failure("MEMBER_NOT_FOUND", $"'{lastSeg}' が見つかりません");

                object? current;
                try { current = prop != null ? prop.GetValue(parent) : field!.GetValue(parent); }
                catch (Exception ex) { return Failure("REFLECTION_READ_FAILED", ex.InnerException?.Message ?? ex.Message); }
                var memberType = prop?.PropertyType ?? field!.FieldType;
                var vProp = current?.GetType().GetProperty("Value");
                if (current != null && current.GetType().Name.Contains("ReactiveProperty") && vProp?.CanWrite == true)
                    return WriteReflectionValue(valElem, vProp.PropertyType,
                        value => vProp.SetValue(current, value), () => vProp.GetValue(current), target, path);
                if ((prop != null && !prop.CanWrite) || field?.IsInitOnly == true)
                    return Failure("MEMBER_READ_ONLY", $"'{lastSeg}' は書き込み不可");
                return WriteReflectionValue(valElem, memberType,
                    value => { if (prop != null) prop.SetValue(parent, value); else field!.SetValue(parent, value); },
                    () => prop != null ? prop.GetValue(parent) : field!.GetValue(parent), target, path);
            });
        }

        private static object WriteReflectionValue(JsonElement input, Type valueType,
            Action<object?> write, Func<object?> read, string target, string path)
        {
            object? converted;
            try
            {
                converted = ConvertJson(input, valueType);
                var underlying = Nullable.GetUnderlyingType(valueType) ?? valueType;
                if (converted == null ? valueType.IsValueType && Nullable.GetUnderlyingType(valueType) == null
                    : !underlying.IsInstanceOfType(converted))
                    return Failure("INVALID_ARGUMENT", $"value を {valueType.Name} に変換できません");
            }
            catch (Exception ex) { return Failure("INVALID_ARGUMENT", ex.InnerException?.Message ?? ex.Message); }
            try { write(converted); }
            catch (Exception ex) { return Failure("REFLECTION_SET_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
            object? actual;
            try { actual = read(); }
            catch (Exception ex) { return Failure("REFLECTION_VERIFY_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
            if (!Equals(converted, actual))
                return new { success = false, error_code = "REFLECTION_VERIFY_FAILED",
                    error = "設定後の値が要求と一致しません", retryable = false, outcome_unknown = true,
                    target, path, expected = ToJsonSafe(converted), actual = ToJsonSafe(actual) };
            return new { success = true, target, path, valueType = valueType.Name,
                verified = true, value = ToJsonSafe(actual) };
        }

        /// <summary>任意のオブジェクトの任意メソッドを引数付きで呼び出す（戻り値がTaskならawait）。</summary>
        private async Task<object> ReflectInvoke(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string target = GetStr(b, "target", "Main");
            string method = GetStr(b, "method", "");
            if (string.IsNullOrEmpty(method)) return Failure("INVALID_ARGUMENT", "method パラメータが必要です");

            JsonElement[] argElems = System.Array.Empty<JsonElement>();
            if (b.TryGetValue("args", out var ae) && ae.ValueKind == JsonValueKind.Array)
                argElems = ae.EnumerateArray().ToArray();

            // UIスレッドでメソッドを解決して呼び出し（Taskなら後でawait）
            var (task, immediate, errorCode, err) = Application.Current.Dispatcher.Invoke(() =>
            {
                var obj = ResolveTarget(target);
                if (obj == null) return ((Task?)null, (object?)null, "TARGET_NOT_FOUND", $"target '{target}' 取得失敗");

                var candidates = obj.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => m.Name == method && m.GetParameters().Length == argElems.Length).ToArray();
                if (candidates.Length == 0)
                {
                    var avail = obj.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Where(m => m.Name == method).Select(m => m.Name + "(" + string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name)) + ")").ToArray();
                    return ((Task?)null, (object?)null, "METHOD_NOT_FOUND", $"メソッド '{method}'({argElems.Length}引数) 見つからず。候補: {string.Join("; ", avail)}");
                }

                var m2 = candidates[0];
                var ps = m2.GetParameters();
                var args = new object?[ps.Length];
                try
                {
                    for (int i = 0; i < ps.Length; i++) args[i] = ConvertJson(argElems[i], ps[i].ParameterType);
                }
                catch (Exception ex) { return ((Task?)null, (object?)null, "INVALID_ARGUMENT", ex.InnerException?.Message ?? ex.Message); }

                try
                {
                    var ret = m2.Invoke(obj, args);
                    if (ret is Task t) return (t, (object?)null, "", "");
                    return ((Task?)null, (object?)ToJsonSafe(ret), "", "");
                }
                catch (Exception ex) { return ((Task?)null, (object?)null, "REFLECTION_INVOKE_FAILED", ex.InnerException?.Message ?? ex.Message); }
            });

            if (!string.IsNullOrEmpty(err)) return Failure(errorCode, err, errorCode == "REFLECTION_INVOKE_FAILED");
            if (task != null)
            {
                try { await task; }
                catch (Exception ex) { return Failure("REFLECTION_INVOKE_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
                // Task<T> なら結果を取得
                var resultProp = task.GetType().GetProperty("Result");
                object? r = null;
                try { if (resultProp != null && task.GetType().IsGenericType) r = ToJsonSafe(resultProp.GetValue(task)); } catch { }
                return new { success = true, target, method, result = r };
            }
            return new { success = true, target, method, result = immediate };
        }

        /// <summary>オブジェクトの型・プロパティ・メソッド・コマンドを一覧する（機能の発見用）。</summary>
        private object InspectObject(HttpListenerRequest req)
        {
            string target = req.QueryString["target"] ?? "Main";
            string path = req.QueryString["path"] ?? "";
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var baseObj = ResolveTarget(target);
                if (baseObj == null) return Failure("TARGET_NOT_FOUND", $"target '{target}' 取得失敗");
                var obj = string.IsNullOrEmpty(path) ? baseObj : ResolvePath(baseObj, path);
                if (obj == null) return Failure("PATH_NOT_FOUND", $"path '{path}' 取得失敗");

                var t = obj.GetType();
                var props = t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(p => p.GetIndexParameters().Length == 0)
                    .OrderBy(p => p.Name)
                    .Select(p =>
                    {
                        object? v = null; try { v = p.GetValue(obj); } catch { }
                        return new
                        {
                            name = p.Name,
                            type = p.PropertyType.Name,
                            canWrite = p.CanWrite,
                            isCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType),
                            value = ToJsonSafe(v)
                        };
                    }).ToList();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => new { name = m.Name, sig = "(" + string.Join(",", m.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}")) + ")", returns = m.ReturnType.Name })
                    .GroupBy(m => m.name).Select(g => g.First()).OrderBy(m => m.name).ToList();
                return (object)new { success = true, target, path, type = t.FullName, properties = props, methods };
            });
        }

        /// <summary>値をJSONシリアライズ可能な形に安全に変換する（複雑なオブジェクトは型名と文字列表現）。</summary>
        private static object? ToJsonSafe(object? val)
        {
            if (val == null) return null;
            var t = val.GetType();
            if (t.IsPrimitive || val is string || val is decimal) return val;
            if (val is Enum) return val.ToString();
            // ReactiveProperty の .Value を展開
            if (t.Name.Contains("ReactiveProperty"))
            {
                var vProp = t.GetProperty("Value");
                if (vProp != null) return ToJsonSafe(vProp.GetValue(val));
            }
            // コレクションは件数と最初の数件
            if (val is System.Collections.IEnumerable en && val is not string)
            {
                var list = new List<object?>();
                int n = 0;
                foreach (var e in en) { if (n++ >= 20) break; list.Add(e?.GetType().Name == null ? null : (e.GetType().IsPrimitive || e is string ? e : e.ToString())); }
                return new { count = n, items = list };
            }
            return val.ToString();
        }

        /// <summary>JsonElement を指定の .NET 型に変換する。</summary>
        private static object? ConvertJson(JsonElement el, Type targetType)
        {
            var underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;
            try
            {
                if (underlying.IsEnum)
                {
                    if (el.ValueKind == JsonValueKind.String) return Enum.Parse(underlying, el.GetString()!, true);
                    if (el.ValueKind == JsonValueKind.Number) return Enum.ToObject(underlying, el.GetInt32());
                }
                return el.ValueKind switch
                {
                    JsonValueKind.String => underlying == typeof(string) ? el.GetString() : Convert.ChangeType(el.GetString(), underlying),
                    JsonValueKind.Number => Convert.ChangeType(el.GetDouble(), underlying),
                    JsonValueKind.True or JsonValueKind.False => el.GetBoolean(),
                    JsonValueKind.Null => null,
                    _ => el.ToString()
                };
            }
            catch { return el.ToString(); }
        }

        // ── デバッグ ──────────────────────────────────────────

        private object DebugViewModel()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var props = vm.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name, isCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType) })
                    .OrderBy(p => p.name).ToArray();
                return (object)new { vmTypeName = vm.GetType().FullName, properties = props };
            });
        }

        private object DebugTimeline()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var props = tvm.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                return (object)new { timelineVmType = tvm.GetType().Name, properties = props };
            });
        }

        private object DebugItems()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var rawItems = GetPropEnum(tvm, "Items");
                if (rawItems == null) return (object)new { error = "Items取得失敗" };
                var result = new List<object>();
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    var props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => { object? v = null; try { v = p.GetValue(item)?.ToString(); } catch { } return new { name = p.Name, value = v }; })
                        .Where(p => p.value != null).ToArray();
                    result.Add(new { typeName = item.GetType().Name, properties = props });
                    break;
                }
                return (object)result;
            });
        }

        private object DebugScene()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var docVMs = GetPropEnum(vm, "DocumentViewModels");
                if (docVMs == null) return (object)new { error = "DocumentViewModels取得失敗" };
                var result = new List<object>();
                foreach (var docVM in docVMs)
                {
                    var sceneVM = GetPropObj(docVM, "ViewModel");
                    result.Add(new
                    {
                        docVMType = docVM.GetType().Name,
                        sceneVMType = sceneVM?.GetType().Name,
                        sceneVMProps = sceneVM?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    });
                    break;
                }
                return (object)result;
            });
        }

        private object DebugVoiceTypes()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                // ロード済みアセンブリからVoiceItem系のクラスを探す
                var voiceTypes = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return Array.Empty<Type>(); } })
                    .Where(t => t.Name.Contains("Voice") && t.Name.Contains("Item") && !t.IsInterface && !t.IsAbstract)
                    .Select(t => new { fullName = t.FullName, ctors = t.GetConstructors().Select(c => string.Join(", ", c.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}"))).ToArray() })
                    .ToArray();

                return (object)new { voiceTypes };
            });
        }

        private object DebugTimelineMethods()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                // timelineフィールドのメソッドを確認
                var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                var timelineObj = timelineField?.GetValue(tvm);
                var sceneField = tvm.GetType().GetField("scene", BindingFlags.NonPublic | BindingFlags.Instance);
                var sceneObj = sceneField?.GetValue(tvm);

                var timelineMethods = timelineObj?.GetType()
                    .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name))})")
                    .OrderBy(s => s).ToArray();

                var sceneMethods = sceneObj?.GetType()
                    .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name))})")
                    .OrderBy(s => s).ToArray();

                return (object)new { timelineMethods, sceneMethods };
            });
        }

        private object DebugMenuItem()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                var menuVM = GetPropObj(tvm, "AddVoiceItemTemplateContextMenuViewModel");
                var menuItems = menuVM?.GetType().GetProperty("Items")?.GetValue(menuVM) as System.Collections.IEnumerable;

                // 再帰的にコマンドを探す
                var result = new List<object>();
                void Explore(object item, int depth)
                {
                    if (depth > 4) return;
                    var t = item.GetType();
                    var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name, IsCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType) })
                        .ToArray();
                    result.Add(new { depth, itemType = t.Name, props });

                    // 子Items再帰
                    var childItems = t.GetProperty("Items")?.GetValue(item) as System.Collections.IEnumerable;
                    if (childItems != null)
                        foreach (var child in childItems) { Explore(child, depth + 1); break; } // 1件のみ
                }
                if (menuItems != null)
                    foreach (var item in menuItems) { Explore(item, 0); break; }

                return (object)result;
            });
        }

        private object DebugVoiceCmd()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                // AddVoiceItemCommandParameterの中身
                var paramProp = tvm.GetType().GetProperty("AddVoiceItemCommandParameter");
                var paramVal = paramProp?.GetValue(tvm);
                var innerVal = paramVal?.GetType().GetProperty("Value")?.GetValue(paramVal);

                // AddVoiceItemTemplateContextMenuViewModel.Items の中身
                var menuVM = GetPropObj(tvm, "AddVoiceItemTemplateContextMenuViewModel");
                object[]? menuItems = null;
                if (menuVM != null)
                {
                    var items = menuVM.GetType().GetProperty("Items")?.GetValue(menuVM) as System.Collections.IEnumerable;
                    menuItems = items?.Cast<object>().Select(item => new
                    {
                        type = item.GetType().Name,
                        props = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray()
                    } as object).ToArray();
                }

                // CurrentCharacterの中身
                var charProp = tvm.GetType().GetProperty("CurrentCharacter");
                var charVal = charProp?.GetValue(tvm);
                var charInner = charVal?.GetType().GetProperty("Value")?.GetValue(charVal);

                // Characters一覧
                var charsEnum = GetPropEnum(tvm, "Characters");
                var charNames = charsEnum?.Cast<object>().Select(c =>
                {
                    var t = c.GetType();
                    return t.GetProperties().Where(p => p.CanRead).Select(p =>
                    {
                        string? v = null; try { v = p.GetValue(c)?.ToString(); } catch { }
                        return new { p.Name, val = v };
                    }).ToArray();
                }).ToArray();

                return (object)new
                {
                    paramType = paramVal?.GetType().Name,
                    innerValType = innerVal?.GetType().Name,
                    innerValProps = innerVal?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    menuItems,
                    charInnerType = charInner?.GetType().Name,
                    charNames
                };
            });
        }

        private object DebugItemToolBar()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                var tvm = GetPropObj(vm!, "ActiveTimelineViewModel");
                var toolbar = GetPropObj(tvm!, "ToolBar");
                var itemToolBar = GetPropObj(toolbar!, "ItemToolBar");
                if (itemToolBar == null) return (object)new { error = "ItemToolBar取得失敗" };

                var props = itemToolBar.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                return (object)new { type = itemToolBar.GetType().FullName, props };
            });
        }

        private object DebugToolBar()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };
                var toolbar = GetPropObj(tvm, "ToolBar");
                if (toolbar == null) return (object)new { error = "ToolBar取得失敗" };

                var props = toolbar.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new { name = p.Name, typeName = p.PropertyType.Name }).OrderBy(p => p.name).ToArray();
                var methods = toolbar.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName).Select(m => m.Name).OrderBy(n => n).ToArray();
                return (object)new { toolbarType = toolbar.GetType().FullName, props, methods };
            });
        }

        private object DebugSceneFields()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { error = "TimelineVM取得失敗" };

                var sceneField = tvm.GetType().GetField("scene", BindingFlags.NonPublic | BindingFlags.Instance);
                var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                var sceneObj = sceneField?.GetValue(tvm);
                var timelineObj = timelineField?.GetValue(tvm);

                return (object)new
                {
                    sceneType = sceneObj?.GetType().FullName,
                    sceneProps = sceneObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                    timelineType = timelineObj?.GetType().FullName,
                    timelineProps = timelineObj?.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new { p.Name, TypeName = p.PropertyType.Name }).ToArray(),
                };
            });
        }

        // ── ヘルパー ──────────────────────────────────────────

        private static object? GetMainViewModel()
        {
            var application = Application.Current;
            if (application == null) return null;
            return MainViewModelSelector.Select(application.Windows.OfType<Window>()
                .Select(w => (Context: w.DataContext, Active: w.IsActive, Visible: w.IsVisible)));
        }

        private static object? GetPropObj(object o, string n) { try { var t = o.GetType(); return (t.GetProperty(n) ?? t.GetProperty(n, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))?.GetValue(o) ?? (t.GetField(n) ?? t.GetField(n, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))?.GetValue(o); } catch { return null; } }
        private static string GetPropStr(object o, string n) { try { return (o.GetType().GetProperty(n) ?? o.GetType().GetProperty(n, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))?.GetValue(o) as string ?? ""; } catch { return ""; } }
        private static System.Collections.IEnumerable? GetPropEnum(object o, string n) { try { return o.GetType().GetProperty(n)?.GetValue(o) as System.Collections.IEnumerable; } catch { return null; } }

        private object GetProps(HttpListenerRequest req)
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var results = new Dictionary<string, object>();
                results["MainViewModel"] = GetObjProps(vm);

                var player = GetPropObj(vm, "PlayerViewModel") ?? GetPropObj(vm, "Player") ?? GetPropObj(vm, "player") ?? GetPropObj(vm, "_player");
                if (player != null) results["PlayerViewModel"] = GetObjProps(player);

                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm != null) results["ActiveTimelineViewModel"] = GetObjProps(tvm);

                var project = GetPropObj(vm, "Project") ?? GetPropObj(vm, "project") ?? GetPropObj(vm, "_project");
                if (project != null) results["Project"] = GetObjProps(project);

                return (object)results;
            });
        }

        private object SearchProps(HttpListenerRequest req)
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                var results = new List<string>();
                SearchPropsRecursive(vm, "Main", results, 0);
                return results;
            });
        }

        private object GetTypeInfo(HttpListenerRequest req)
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };

                var targetName = req.QueryString["name"] ?? "Project";
                object? obj = targetName switch
                {
                    "Project" => GetPropObj(vm, "Project") ?? GetPropObj(vm, "project") ?? GetPropObj(vm, "_project"),
                    "Player" => GetPropObj(vm, "PlayerViewModel") ?? GetPropObj(vm, "Player") ?? GetPropObj(vm, "player") ?? GetPropObj(vm, "_player"),
                    "ActiveTimeline" => GetPropObj(vm, "ActiveTimelineViewModel"),
                    _ => null
                };

                if (obj == null) return (object)new { error = $"{targetName} not found" };

                var type = obj.GetType();
                var bases = new List<string>();
                var t = type.BaseType;
                while (t != null) { bases.Add(t.FullName ?? t.Name); t = t.BaseType; }

                return (object)new
                {
                    name = targetName,
                    fullType = type.FullName,
                    baseTypes = bases,
                    interfaces = type.GetInterfaces().Select(i => i.FullName).ToList(),
                    members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance).Select(m => $"{m.MemberType}: {m.Name}").ToList()
                };
            });
        }

        private void SearchPropsRecursive(object o, string path, List<string> results, int depth)
        {
            if (depth > 5 || o == null) return;
            var type = o.GetType();
            if (type.IsPrimitive || type == typeof(string) || type == typeof(TimeSpan) || type == typeof(DateTime) || type == typeof(Guid)) return;

            // プロパティ
            foreach (var p in type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                try
                {
                    var name = p.Name;
                    if (p.GetIndexParameters().Length > 0) continue;

                    var val = p.GetValue(o);
                    if (val == null) continue;

                    string valStr = val.ToString() ?? "null";
                    bool isReactive = val.GetType().Name.Contains("ReactiveProperty");
                    if (isReactive)
                    {
                        var vProp = val.GetType().GetProperty("Value");
                        if (vProp != null) valStr = $"{valStr} (Value: {vProp.GetValue(val)?.ToString() ?? "null"})";
                    }

                    bool match = name.Contains("Time") || name.Contains("Frame") || name.Contains("Position") ||
                                 name.Contains("Player") || name.Contains("Preview") || name.Contains("Scene") ||
                                 name.Contains("Project") || name.Contains("Seek") || name.Contains("Playback");

                    if (match)
                    {
                        results.Add($"{path}.{name} (Prop:{p.PropertyType.Name}): {valStr}");
                    }

                    if (!type.Assembly.FullName.Contains("System") && depth < 5)
                    {
                        SearchPropsRecursive(val, $"{path}.{name}", results, depth + 1);
                    }
                }
                catch { }
            }

            // フィールド
            foreach (var f in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                try
                {
                    var name = f.Name;
                    var val = f.GetValue(o);
                    if (val == null) continue;

                    string valStr = val.ToString() ?? "null";
                    bool isReactive = val.GetType().Name.Contains("ReactiveProperty");
                    if (isReactive)
                    {
                        var vProp = val.GetType().GetProperty("Value");
                        if (vProp != null) valStr = $"{valStr} (Value: {vProp.GetValue(val)?.ToString() ?? "null"})";
                    }

                    bool match = name.Contains("Time") || name.Contains("Frame") || name.Contains("Position") ||
                                 name.Contains("Player") || name.Contains("Preview") || name.Contains("Scene") ||
                                 name.Contains("Project") || name.Contains("Seek") || name.Contains("Playback") ||
                                 name.Contains("project") || name.Contains("player") || name.Contains("scene");

                    if (match)
                    {
                        results.Add($"{path}.{name} (Field:{f.FieldType.Name}): {valStr}");
                    }

                    if (!type.Assembly.FullName.Contains("System") && depth < 5)
                    {
                        SearchPropsRecursive(val, $"{path}.{name}", results, depth + 1);
                    }
                }
                catch { }
            }
        }

        private static object GetObjProps(object o)
        {
            var props = new List<object>();
            var type = o.GetType();

            while (type != null && type != typeof(object))
            {
                // プロパティ
                foreach (var p in type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                {
                    if (p.GetIndexParameters().Length > 0) continue;
                    object? val = null;
                    try { val = p.GetValue(o); } catch { }
                    string typeName = p.PropertyType.Name;
                    string valStr = val?.ToString() ?? "null";

                    if (val != null && val.GetType().Name.Contains("ReactiveProperty"))
                    {
                        try
                        {
                            var vProp = val.GetType().GetProperty("Value");
                            if (vProp != null) valStr = $"{valStr} (Value: {vProp.GetValue(val)?.ToString() ?? "null"})";
                        }
                        catch { }
                    }

                    props.Add(new { name = p.Name, type = typeName, value = valStr, isField = false, declType = type.Name });
                }

                // フィールド
                foreach (var f in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                {
                    object? val = null;
                    try { val = f.GetValue(o); } catch { }
                    string typeName = f.FieldType.Name;
                    string valStr = val?.ToString() ?? "null";

                    if (val != null && val.GetType().Name.Contains("ReactiveProperty"))
                    {
                        try
                        {
                            var vProp = val.GetType().GetProperty("Value");
                            if (vProp != null) valStr = $"{valStr} (Value: {vProp.GetValue(val)?.ToString() ?? "null"})";
                        }
                        catch { }
                    }

                    props.Add(new { name = f.Name, type = typeName, value = valStr, isField = true, declType = type.Name });
                }
                type = type.BaseType!;
            }

            return props;
        }

        private static object? GetPropValue(object? o, string n)
        {
            if (o == null) return null;
            var prop = o.GetType().GetProperty(n);
            if (prop == null) return null;
            var val = prop.GetValue(o);
            if (val == null) return null;
            // ReactiveProperty なら .Value を取得
            var vProp = val.GetType().GetProperty("Value");
            if (vProp != null) return vProp.GetValue(val);
            return val;
        }

        private static bool SetPropValue(object? o, string n, object v)
        {
            if (o == null) return false;
            var prop = o.GetType().GetProperty(n);
            if (prop == null) return false;
            var val = prop.GetValue(o);
            // ReactiveProperty なら .Value に設定
            var vProp = val?.GetType().GetProperty("Value");
            if (vProp != null && vProp.CanWrite)
            {
                try { vProp.SetValue(val, Convert.ChangeType(v, vProp.PropertyType)); return true; } catch { }
            }
            if (prop.CanWrite)
            {
                try { prop.SetValue(o, Convert.ChangeType(v, prop.PropertyType)); return true; } catch { }
            }
            return false;
        }

        private static bool TryCmd(object vm, string name, object? param = null)
        {
            try { if (vm.GetType().GetProperty(name)?.GetValue(vm) is ICommand c && c.CanExecute(param)) { c.Execute(param); return true; } return false; }
            catch { return false; }
        }

        private static bool TryMethod(object vm, string name, params object[] args)
        {
            try { var m = vm.GetType().GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); if (m == null) return false; m.Invoke(vm, args.Take(m.GetParameters().Length).ToArray<object?>()); return true; }
            catch { return false; }
        }

        // PreviewViewModel を使った再生/停止
        private static async Task<object> PlaybackControl(string action)
        {
            try
            {
                var preview = Application.Current.Dispatcher.Invoke(() => GetPreviewViewModel());
                if (preview == null) return new { success = false, error = "PreviewViewModel not found" };
                if (action == "play")
                    await InvokeAsyncMethod(preview, "TogglePlayAsync");
                else
                    await InvokeAsyncMethod(preview, "StopAsync");
                return new { success = true, action };
            }
            catch (Exception ex) { return new { success = false, error = ex.Message }; }
        }

        // AnchorableAreaViewModels から PreviewViewModel を取得
        private static object? GetPreviewViewModel()
        {
            var vm = GetMainViewModel();
            if (vm == null) return null;
            var areas = GetPropEnum(vm, "AnchorableAreaViewModels");
            if (areas == null) return null;
            foreach (var area in areas)
            {
                var innerVm = GetPropObj(area, "ViewModel") ?? area;
                if (innerVm.GetType().Name == "PreviewViewModel") return innerVm;
            }
            return null;
        }

        // PreviewViewModel のメソッドを async で呼び出す（Task を await）
        private static async Task InvokeAsyncMethod(object target, string methodName, params object[] args)
        {
            // オーバーロードがある場合は引数型で一致するものを選ぶ
            var methods = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(m => m.Name == methodName).ToArray();
            System.Reflection.MethodInfo? m2 = null;
            if (args.Length > 0)
                m2 = methods.FirstOrDefault(m => m.GetParameters().Length == args.Length &&
                     m.GetParameters()[0].ParameterType.IsAssignableFrom(args[0].GetType()));
            m2 ??= methods.FirstOrDefault(m => m.GetParameters().Length == args.Length);
            m2 ??= methods.FirstOrDefault();
            if (m2 == null) throw new InvalidOperationException($"Method '{methodName}' not found on {target.GetType().Name}");
            var result = m2.Invoke(target, args.Take(m2.GetParameters().Length).ToArray<object?>());
            if (result is Task t) await t;
        }

        /// <summary>
        /// UIスレッド上で非同期メソッドを実行し、内側の Task 完了まで待つ。
        /// Dispatcher.InvokeAsync(async () => ...) だと Operation 完了と内側 Task 完了が別になる。
        /// </summary>
        private static Task RunOnUi(Func<Task> action)
        {
            var dispatcher = Application.Current?.Dispatcher
                ?? throw new InvalidOperationException("WPF Dispatcher unavailable");
            if (dispatcher.CheckAccess()) return action();
            return dispatcher.InvokeAsync(action).Task.Unwrap();
        }

        private static Task<T> RunOnUi<T>(Func<Task<T>> action)
        {
            var dispatcher = Application.Current?.Dispatcher
                ?? throw new InvalidOperationException("WPF Dispatcher unavailable");
            if (dispatcher.CheckAccess()) return action();
            return dispatcher.InvokeAsync(action).Task.Unwrap();
        }

        private static async Task<Dictionary<string, JsonElement>> ReadBody(HttpListenerRequest req)
        {
            const int maxBytes = 1024 * 1024;
            if (req.ContentLength64 > maxBytes) throw new ArgumentException("Request body exceeds 1 MiB");
            using var buffer = new MemoryStream();
            var chunk = new byte[8192];
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            int count;
            while ((count = await req.InputStream.ReadAsync(chunk.AsMemory(), deadline.Token)) > 0)
            {
                if (buffer.Length + count > maxBytes) throw new ArgumentException("Request body exceeds 1 MiB");
                buffer.Write(chunk, 0, count);
            }
            if (buffer.Length == 0) return new();
            return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(buffer.ToArray())
                ?? throw new ArgumentException("JSON object required");
        }

        private static string GetStr(Dictionary<string, JsonElement> d, string k, string def) => d.TryGetValue(k, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? def : def;
        private static int GetInt(Dictionary<string, JsonElement> d, string k, int def)
        {
            return TimelineInputValidation.GetInt(d, k, def);
        }
        private static double GetDouble(Dictionary<string, JsonElement> d, string k, double def)
        {
            if (!d.TryGetValue(k, out var v)) return def;
            if (v.ValueKind == JsonValueKind.Number)
            {
                if (v.TryGetDouble(out double number) && double.IsFinite(number)) return number;
            }
            else if (v.ValueKind == JsonValueKind.String
                     && double.TryParse(v.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                     && double.IsFinite(parsed))
            {
                return parsed;
            }
            throw new ArgumentException(k + " must be a finite number");
        }
        private void Log(string msg) => LogMessage?.Invoke($"[{DateTime.Now:HH:mm:ss}] {msg}");

        // ================================================================
        // PlayerViewModel のプロパティ・コマンド・メソッドを軽量調査
        private static object DebugPlayer()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "VM失敗" };

                // AnchorableAreaViewModels の中からPlayerっぽいものを探す
                var areas = GetPropEnum(vm, "AnchorableAreaViewModels");
                var found = new List<object>();
                if (areas != null)
                {
                    foreach (var area in areas)
                    {
                        var innerVm = GetPropObj(area, "ViewModel") ?? area;
                        var typeName = innerVm.GetType().Name;
                        var props = innerVm.GetType()
                            .GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                            .Select(p => {
                                string? val = null;
                                try { var v = p.GetValue(innerVm); val = v?.GetType().GetProperty("Value")?.GetValue(v)?.ToString() ?? v?.ToString(); } catch { }
                                return new { p.Name, type = p.PropertyType.Name, val };
                            }).ToArray();
                        var commands = innerVm.GetType()
                            .GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                            .Where(p => typeof(System.Windows.Input.ICommand).IsAssignableFrom(p.PropertyType))
                            .Select(p => p.Name).ToArray();
                        var methods = innerVm.GetType()
                            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                            .Where(m => !m.IsSpecialName)
                            .Select(m => $"{m.Name}({string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name))})")
                            .ToArray();
                        found.Add(new { areaType = area.GetType().Name, vmType = typeName, props, commands, methods });
                    }
                }
                return (object)new { count = found.Count, areas = found };
            });
        }

        // VisualTree全列挙（プレビュー要素名調査用）
        private static object DebugVisualTree(HttpListenerRequest req)
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var window = Application.Current.MainWindow;
                if (window == null) return (object)new { error = "MainWindow not found" };
                var results = new List<object>();
                var queue = new Queue<(DependencyObject obj, int depth)>();
                queue.Enqueue((window, 0));
                int count = 0;
                while (queue.Count > 0 && count < 500)
                {
                    var (cur, depth) = queue.Dequeue();
                    if (cur is FrameworkElement fe)
                    {
                        double w = fe.ActualWidth, h = fe.ActualHeight;
                        if (w > 50 || h > 50 || !string.IsNullOrEmpty(fe.Name))
                        {
                            results.Add(new {
                                depth,
                                name = fe.Name,
                                type = fe.GetType().Name,
                                w = Math.Round(w), h = Math.Round(h)
                            });
                            count++;
                        }
                    }
                    int children = VisualTreeHelper.GetChildrenCount(cur);
                    for (int i = 0; i < children; i++) queue.Enqueue((VisualTreeHelper.GetChild(cur, i), depth + 1));
                }
                return (object)new { count = results.Count, elements = results };
            });
        }

        // VisualTreeをBFS探索（Name または 型名で検索）
        private static UIElement? FindVisualByName(DependencyObject root, string name)
        {
            if (root == null) return null;
            var queue = new Queue<DependencyObject>();
            queue.Enqueue(root);
            while (queue.Count > 0)
            {
                var cur = queue.Dequeue();
                if (cur is FrameworkElement fe && cur is UIElement ui)
                {
                    // Name プロパティ または 型名で一致
                    if (fe.Name == name || fe.GetType().Name == name)
                        return ui;
                }
                int count = VisualTreeHelper.GetChildrenCount(cur);
                for (int i = 0; i < count; i++) queue.Enqueue(VisualTreeHelper.GetChild(cur, i));
            }
            return null;
        }

    }
}
