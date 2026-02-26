using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
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
    public class McpHttpServer
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

        private HttpListener? _listener;
        private CancellationTokenSource? _cts;
        private bool _isRunning = false;

        public const int Port = 8765;
        public string BaseUrl => $"http://localhost:{Port}/";
        public bool IsRunning => _isRunning;
        public event Action<string>? LogMessage;

        public void Start()
        {
            if (_isRunning) return;
            _cts = new CancellationTokenSource();
            _listener = new HttpListener();
            _listener.Prefixes.Add(BaseUrl);
            _listener.Start();
            _isRunning = true;
            Task.Run(() => ListenLoop(_cts.Token));
            Log($"MCPサーバー起動: {BaseUrl}");
        }

        public void Stop()
        {
            if (!_isRunning) return;
            _cts?.Cancel();
            _listener?.Stop();
            _listener = null;
            _isRunning = false;
            Log("MCPサーバー停止");
        }

        private async Task ListenLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested && _listener != null)
            {
                try { var ctx = await _listener.GetContextAsync(); _ = Task.Run(() => HandleRequest(ctx), ct); }
                catch (HttpListenerException) { break; }
                catch (ObjectDisposedException) { break; }
                catch (Exception ex) { Log($"エラー: {ex.Message}"); }
            }
        }

        private async Task HandleRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;
            res.AddHeader("Access-Control-Allow-Origin", "*");
            res.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            res.AddHeader("Access-Control-Allow-Headers", "Content-Type");
            if (req.HttpMethod == "OPTIONS") { res.StatusCode = 204; res.Close(); return; }

            try
            {
                string path = req.Url?.AbsolutePath ?? "/";
                Log($"{req.HttpMethod} {path}");
                object? result = (req.HttpMethod, path) switch
                {
                    ("GET", "/api/status") => GetStatus(),
                    ("GET", "/api/project") => GetProjectInfo(),
                    ("GET", "/api/items") => GetTimelineItems(),
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
                    ("POST", "/api/items/text") => await AddTextItem(req),
                    ("POST", "/api/items/voice") => await AddVoiceItem(req),
                    ("POST", "/api/items/reorder") => await ReorderItems(req),
                    ("POST", "/api/items/arrange") => await ArrangeItems(req),
                    ("POST", "/api/items/move") => await MoveItem(req),
                    ("POST", "/api/items/tachie") => await AddTachieItem(req),
                    ("POST", "/api/items/face") => await AddFaceItem(req),
                    ("POST", "/api/items/face/param") => await SetFaceParam(req),
                    ("POST", "/api/items/effect/video") => await AddVideoEffect(req),
                    ("POST", "/api/items/effect") => await AddEffectToItem(req),
                    ("POST", "/api/items/prop") => await SetItemProp(req),
                    ("GET", "/api/effects/list") => ListEffects(),
                    ("POST", "/api/items/delete") => await DeleteItems(req),
                    ("GET", "/api/debug/tachie") => DebugTachie(),
                    ("GET", "/api/debug/tachie/props") => DebugTachieItemProps(),
                    ("GET", "/api/debug/facetypes") => DebugFaceTypes(),
                    ("GET", "/api/debug/voiceitem/props") => DebugVoiceItemProps(),
                    ("POST", "/api/items/effect/audio") => await AddAudioEffect(req),
                    ("GET", "/api/debug/visualtree") => DebugVisualTree(req),
                    ("GET", "/api/debug/player") => DebugPlayer(),
                    ("GET", "/api/preview/capture") => CapturePreview(req),
                    ("POST", "/api/preview/seek") => await SeekAndCapture(req),
                    ("GET", "/api/preview/position") => GetPlaybackPosition(),
                    ("POST", "/api/preview/record") => await RecordAudio(req),
                    ("POST", "/api/preview/watch") => await WatchScene(req),
                    ("POST", "/api/timeline/duration") => await SetTimelineDuration(req),
                    ("POST", "/api/playback/play") => await PlaybackControl("play"),
                    ("POST", "/api/playback/stop") => await PlaybackControl("stop"),
                    ("POST", "/api/project/save") => ExecCommand("SaveProjectCommand"),
                    _ => null
                };
                if (result == null) { await WriteJson(res, 404, new { error = "Not Found", path }); return; }
                await WriteJson(res, 200, result);
            }
            catch (Exception ex) { Log($"エラー: {ex.Message}"); await WriteJson(res, 500, new { error = ex.Message }); }
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

        private object GetStatus() => new { status = "running", version = "1.0.0", port = Port, timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") };

        private object GetProjectInfo()
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { error = "MainViewModel取得失敗" };
                return new { vmType = vm.GetType().FullName, projectName = GetPropStr(vm, "ProjectName"), projectPath = GetPropStr(vm, "ProjectPath") };
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
                        var item = GetPropObj(iv, "Item") ?? iv;
                        items.Add(new
                        {
                            layer = GetPropObj(item, "Layer") ?? GetPropObj(iv, "Layer"),
                            frame = GetPropObj(item, "Frame") ?? GetPropObj(iv, "Frame"),
                            length = GetPropObj(item, "Length") ?? GetPropObj(iv, "Length"),
                            type = item.GetType().Name,
                            text = GetPropObj(item, "Serif") ?? GetPropObj(item, "Text")
                                  ?? GetPropObj(item, "FilePath") ?? GetPropObj(item, "Name")
                                  ?? GetPropObj(iv, "Serif") ?? GetPropObj(iv, "Text") ?? GetPropObj(iv, "DisplayName"),
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
            int newFrame = GetInt(b, "frame", 0);
            int newLength = GetInt(b, "length", -1);

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

                    try
                    {
                        var frameProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        var lengthProp = item.GetType().GetProperty("Length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        int oldFrame = (int)(frameProp?.GetValue(item) ?? 0);
                        int oldLength = (int)(lengthProp?.GetValue(item) ?? 0);
                        frameProp?.SetValue(item, newFrame);
                        if (newLength > 0) lengthProp?.SetValue(item, newLength);
                        return (object)new { success = true, filename, oldFrame, oldLength, newFrame, newLength };
                    }
                    catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
                }
                return (object)new { success = false, error = $"ファイルが見つかりません: {filename}" };
            });
        }

        private async Task<object> AddTachieItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string ch = GetStr(b, "character", "ゆっくり霊夢");
            int frame = GetInt(b, "frame", 0);
            int layer = GetInt(b, "layer", 3);
            int length = GetInt(b, "length", 300);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (object)new { success = false, error = "TimelineVM取得失敗" };

                // キャラ検索
                var charsEnum = GetPropEnum(tvm, "Characters");
                object? targetChar = null;
                string searchName = ch.Replace("ゆっくり", "").Trim();
                if (charsEnum != null)
                    foreach (var c in charsEnum)
                    { string? n = null; try { n = c.GetType().GetProperty("Name")?.GetValue(c)?.ToString(); } catch { } if (n != null && n.Contains(searchName)) { targetChar = c; break; } }
                if (targetChar == null) return (object)new { success = false, error = "キャラが見つかりません: " + ch };

                // MainModel.AddTachieItem(int frame, int layer, Character character)
                object? mainModel = GetMainModel(vm);
                if (mainModel == null) return (object)new { success = false, error = "MainModel取得失敗" };

                try
                {
                    var addMethod = mainModel.GetType().GetMethod("AddTachieItem", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (addMethod == null)
                    {
                        var methods = mainModel.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                            .Where(m => m.Name.Contains("Tachie") || m.Name.Contains("Face"))
                            .Select(m => m.Name + "(" + string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name)) + ")")
                            .ToArray();
                        return (object)new { success = false, error = "AddTachieItem見つからず", methods };
                    }
                    addMethod.Invoke(mainModel, new object?[] { frame, layer, targetChar });

                    // 追加されたアイテムのlengthを調整
                    Start_SleepMs(300);
                    var rawItems = GetPropEnum(tvm, "Items");
                    if (rawItems != null)
                        foreach (var iv in rawItems)
                        {
                            var item = GetPropObj(iv, "Item") ?? iv;
                            var fProp = item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                            var lProp = item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                            var lenProp = item.GetType().GetProperty("Length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                            if (fProp == null || lProp == null) continue;
                            try
                            {
                                int f2 = (int)(fProp.GetValue(item) ?? -1);
                                int l2 = (int)(lProp.GetValue(item) ?? -1);
                                if (f2 == frame && l2 == layer && lenProp != null)
                                { lenProp.SetValue(item, length); break; }
                            }
                            catch { }
                        }

                    return (object)new { success = true, character = ch, frame, layer, length };
                }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
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

        // 表情アイテム追加
        private async Task<object> AddFaceItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string ch = GetStr(b, "character", "ゆっくり霊夢");
            int frame = GetInt(b, "frame", 0);
            int layer = GetInt(b, "layer", 4);
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var charsEnum = GetPropEnum(tvm, "Characters");
                object? targetChar = null;
                string sn = ch.Replace("ゆっくり", "").Trim();
                if (charsEnum != null) foreach (var c in charsEnum) { string? n = null; try { n = c.GetType().GetProperty("Name")?.GetValue(c)?.ToString(); } catch { } if (n != null && n.Contains(sn)) { targetChar = c; break; } }
                if (targetChar == null) return (object)new { success = false, error = "キャラなし: " + ch };
                var mm = GetMainModel(vm); if (mm == null) return (object)new { success = false, error = "MainModel失敗" };
                try { var m = mm.GetType().GetMethod("AddFaceItem", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); if (m == null) return (object)new { success = false, error = "AddFaceItem未発見" }; m.Invoke(mm, new object?[] { frame, layer, targetChar }); return (object)new { success = true, character = ch, frame, layer }; }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            });
        }

        // アイテムにVideoEffectを追加
        private async Task<object> AddVideoEffect(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
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
                    return (object)new { success = true, effect = eType.FullName };
                }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            });
        }

        // アイテムの単一プロパティを設定（FadeIn/FadeOutなど）
        private async Task<object> SetItemProp(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int tf = GetInt(b, "frame", 0); int tl = GetInt(b, "layer", 0);
            string pName = GetStr(b, "prop", ""); string pVal = GetStr(b, "value", "");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };
                object? targetItem = null;
                foreach (var iv in rawItems) { var item = GetPropObj(iv, "Item") ?? iv; try { int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); if (f2 == tf && l2 == tl) { targetItem = item; break; } } catch { } }
                if (targetItem == null) return (object)new { success = false, error = $"アイテム未発見 f={tf} l={tl}" };
                var p = targetItem.GetType().GetProperty(pName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (p == null) return (object)new { success = false, error = $"プロパティ'{pName}'なし" };
                try { p.SetValue(targetItem, Convert.ChangeType(pVal, p.PropertyType)); return (object)new { success = true, prop = pName, value = pVal }; }
                catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
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

        // アイテム削除（layer指定 or frame+layer指定）
        private async Task<object> DeleteItems(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            int[]? layers = null;
            if (b.TryGetValue("layers", out var le))
                layers = le.EnumerateArray().Select(x => x.GetInt32()).ToArray();
            int targetFrame = GetInt(b, "frame", -1);
            int targetLayer = GetInt(b, "layer", -1);

            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };

                // 削除対象のItemオブジェクトを収集
                var toRemove = new List<object>();
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    try
                    {
                        int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1);
                        int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1);
                        if (layers != null && layers.Contains(l2)) toRemove.Add(item);
                        else if (targetFrame >= 0 && targetLayer >= 0 && f2 == targetFrame && l2 == targetLayer) toRemove.Add(item);
                    }
                    catch { }
                }
                if (toRemove.Count == 0) return (object)new { success = true, removed = 0, note = "対象アイテムなし" };

                // timelineフィールドからTimelineオブジェクトを取得
                var timelineField = tvm.GetType().GetField("timeline", BindingFlags.NonPublic | BindingFlags.Instance);
                var timelineObj = timelineField?.GetValue(tvm);
                if (timelineObj == null) return (object)new { success = false, error = "timelineフィールド取得失敗" };

                var deleteMethod = timelineObj.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "DeleteItems" && m.GetParameters().Length == 1);
                if (deleteMethod == null) return (object)new { success = false, error = "DeleteItems未発見" };
                try
                {
                    // IEnumerable<IItem> に変換
                    var iItemType = toRemove[0].GetType().GetInterfaces().FirstOrDefault(i => i.Name == "IItem");
                    Array arr;
                    if (iItemType != null)
                    {
                        arr = Array.CreateInstance(iItemType, toRemove.Count);
                        for (int i = 0; i < toRemove.Count; i++) arr.SetValue(toRemove[i], i);
                    }
                    else { arr = toRemove.ToArray(); }
                    deleteMethod.Invoke(timelineObj, new object[] { arr });
                    return (object)new { success = true, removed = toRemove.Count };
                }
                catch (Exception ex) { return (object)new { success = false, error = ex.InnerException?.Message ?? ex.Message }; }
            });
        }

        private async Task<object> AddEffectToItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
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
                    return (object)new { success = true, effect = effectType.Name, item = targetItem.GetType().Name, effectsType = effects?.GetType().FullName };
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
            int tf = GetInt(b, "frame", 0); int tl = GetInt(b, "layer", 0);
            // keyValuePairs: { "FacePath": "...", etc. }
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel(); if (vm == null) return (object)new { success = false, error = "VM失敗" };
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel"); if (tvm == null) return (object)new { success = false, error = "TVM失敗" };
                var rawItems = GetPropEnum(tvm, "Items"); if (rawItems == null) return (object)new { success = false, error = "Items失敗" };
                object? targetItem = null;
                foreach (var iv in rawItems)
                {
                    var item = GetPropObj(iv, "Item") ?? iv;
                    try { int f2 = (int)(item.GetType().GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); int l2 = (int)(item.GetType().GetProperty("Layer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(item) ?? -1); if (f2 == tf && l2 == tl && item.GetType().Name.Contains("Face")) { targetItem = item; break; } } catch { }
                }
                if (targetItem == null) return (object)new { success = false, error = $"FaceItem未発見 f={tf} l={tl}" };
                // FaceParameterを探して設定
                var fpProp = targetItem.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(p => p.Name.Contains("FaceParameter") || p.Name.Contains("FaceParam"));
                var faceParam = fpProp?.GetValue(targetItem);
                var results = new System.Collections.Generic.List<string>();
                foreach (var kv in b)
                {
                    if (kv.Key == "frame" || kv.Key == "layer") continue;
                    // アイテム直接のプロパティ
                    var directProp = targetItem.GetType().GetProperty(kv.Key, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (directProp != null) { try { directProp.SetValue(targetItem, Convert.ChangeType(kv.Value.GetString() ?? kv.Value.ToString(), directProp.PropertyType)); results.Add($"{kv.Key}=OK(item)"); } catch (Exception ex) { results.Add($"{kv.Key}=NG({ex.Message})"); } continue; }
                    // FaceParameter内のプロパティ
                    if (faceParam != null)
                    {
                        var fp2 = faceParam.GetType().GetProperty(kv.Key, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        if (fp2 != null) { try { fp2.SetValue(faceParam, Convert.ChangeType(kv.Value.GetString() ?? kv.Value.ToString(), fp2.PropertyType)); results.Add($"{kv.Key}=OK(faceParam)"); } catch (Exception ex) { results.Add($"{kv.Key}=NG({ex.Message})"); } continue; }
                    }
                    results.Add($"{kv.Key}=NOTFOUND");
                }
                return (object)new { success = true, results };
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
                }
                else return (object)new { success = false, error = "Addメソッドが見つかりません", listType = audioEffects.GetType().FullName };

                return (object)new { success = true, effect = effType.FullName };
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

        private async Task<object> AddTextItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string text = GetStr(b, "text", "テキスト"); int frame = GetInt(b, "frame", 0); int layer = GetInt(b, "layer", 0);
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (object)new { success = false, error = "MainViewModel取得失敗" };
                bool ok = TryMethod(vm, "AddTextItem", text, frame, layer) || TryCmd(vm, "AddTextItemCommand");
                return ok ? (object)new { success = true } : new { success = false, error = "コマンドが見つかりません" };
            });
        }

        private async Task<object> AddVoiceItem(HttpListenerRequest req)
        {
            var b = await ReadBody(req);
            string text = GetStr(b, "text", "セリフ");
            int frame = GetInt(b, "frame", 0);
            int layer = GetInt(b, "layer", 0);
            string ch = GetStr(b, "character", "ゆっくり霊夢");

            // UIスレッドでキャラ・パラメータ取得
            var (targetChar, paramType, paramObj, mainModel, errMsg) = Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return (null, null, null, null, "MainViewModel取得失敗");
                var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
                if (tvm == null) return (null, null, null, null, "TimelineVM取得失敗");

                // キャラクター検索
                var charsEnum = GetPropEnum(tvm, "Characters");
                object? targetChar = null;
                string searchName = ch.Replace("ゆっくり", "").Trim();
                if (charsEnum != null)
                    foreach (var c in charsEnum)
                    { string? n = null; try { n = c.GetType().GetProperty("Name")?.GetValue(c)?.ToString(); } catch { } if (n != null && n.Contains(searchName)) { targetChar = c; break; } }
                if (targetChar == null) return (null, null, null, null, "キャラが見つかりません:" + ch);

                // AddVoiceItemCommandParameter 生成
                var paramType = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                    .FirstOrDefault(t => t.FullName == "YukkuriMovieMaker.ViewModels.CommandParameter.AddVoiceItemCommandParameter");
                if (paramType == null) return (null, null, null, null, "AddVoiceItemCommandParameter型が見つかりません");

                // decorations: IEnumerable<TextDecoration> の空配列を作る
                var decoRP = tvm.GetType().GetProperty("Decorations")?.GetValue(tvm);
                var decoVal = decoRP?.GetType().GetProperty("Value")?.GetValue(decoRP);
                if (decoVal == null)
                {
                    // TextDecoration 型を探して空配列を生成
                    var textDecoType = AppDomain.CurrentDomain.GetAssemblies()
                        .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                        .FirstOrDefault(t => t.Name == "TextDecoration" && t.Namespace?.Contains("YukkuriMovieMaker") == true);
                    decoVal = textDecoType != null
                        ? System.Array.CreateInstance(textDecoType, 0)
                        : (object)System.Array.Empty<object>();
                }

                object? paramObj = null;
                string paramErr = "";
                try { paramObj = Activator.CreateInstance(paramType, new object?[] { frame, layer, targetChar, text, decoVal }); }
                catch (Exception ex) { paramErr = ex.InnerException?.Message ?? ex.Message; }
                if (paramObj == null)
                {
                    // コンストラクタ情報とdecoValの型をデバッグ出力
                    var ctors = paramType.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Select(c => string.Join(",", c.GetParameters().Select(p => p.ParameterType.FullName + " " + p.Name))).ToArray();
                    return (null, null, null, null, $"Param生成失敗: {paramErr} | decoType={decoVal?.GetType().FullName} | ctors={string.Join(";", ctors)}");
                }

                // MainModel 取得
                object? mainModel = null;
                foreach (var f in vm.GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance))
                    if (f.FieldType.Name.Contains("MainModel")) { mainModel = f.GetValue(vm); break; }
                if (mainModel == null)
                    foreach (var p in vm.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
                        if (p.PropertyType.Name.Contains("MainModel")) { try { mainModel = p.GetValue(vm); } catch { } if (mainModel != null) break; }

                return (targetChar, paramType, paramObj, mainModel, "");
            });

            if (errMsg != "") return new { success = false, error = errMsg };
            if (mainModel == null) return new { success = false, error = "MainModel取得失敗" };

            // MainModel.AddVoiceItemAsync(int frame, int layer, Character character, string serif, IEnumerable<TextDecoration> decorations)
            try
            {
                var addMethod = mainModel.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "AddVoiceItemAsync" && m.GetParameters().Length == 5);

                if (addMethod == null)
                    return new { success = false, error = "AddVoiceItemAsync(5args)見つからず" };

                // TextDecoration の空配列
                var textDecoType = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => { try { return a.GetTypes(); } catch { return System.Array.Empty<Type>(); } })
                    .FirstOrDefault(t => t.Name == "TextDecoration" && t.Namespace?.Contains("YukkuriMovieMaker") == true);
                var emptyDecos = textDecoType != null
                    ? (object)System.Array.CreateInstance(textDecoType, 0)
                    : System.Array.Empty<object>();

                var task = Application.Current.Dispatcher.Invoke(() =>
                    addMethod.Invoke(mainModel, new object?[] { frame, layer, targetChar, text, emptyDecos }) as System.Threading.Tasks.Task);
                if (task != null) await task;

                return new { success = true, character = ch, text, frame, layer };
            }
            catch (Exception ex)
            {
                return new { success = false, error = ex.InnerException?.Message ?? ex.Message };
            }
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
            return Application.Current.Windows.OfType<Window>()
                .Select(w => { var dc = w.DataContext; if (dc == null) return (dc, -1); int idx = -1; try { idx = (int)(dc.GetType().GetProperty("Index")?.GetValue(dc) ?? -1); } catch { } return (dc, idx); })
                .Where(x => x.Item2 == 0).Select(x => x.Item1).FirstOrDefault();
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

        private static async Task<Dictionary<string, JsonElement>> ReadBody(HttpListenerRequest req)
        {
            if (req.ContentLength64 <= 0) return new();
            using var r = new System.IO.StreamReader(req.InputStream, Encoding.UTF8);
            var body = await r.ReadToEndAsync();
            return string.IsNullOrWhiteSpace(body) ? new() : JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(body) ?? new();
        }

        private static string GetStr(Dictionary<string, JsonElement> d, string k, string def) => d.TryGetValue(k, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? def : def;
        private static int GetInt(Dictionary<string, JsonElement> d, string k, int def) => d.TryGetValue(k, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : def;
        private void Log(string msg) => LogMessage?.Invoke($"[{DateTime.Now:HH:mm:ss}] {msg}");

        // ================================================================
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
            catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
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
                var buffer = new System.Collections.Concurrent.ConcurrentBag<byte[]>();
                var tcs = new TaskCompletionSource<bool>();

                capture.DataAvailable += (s, e) =>
                {
                    if (e.BytesRecorded > 0)
                    {
                        var chunk = new byte[e.BytesRecorded];
                        Buffer.BlockCopy(e.Buffer, 0, chunk, 0, e.BytesRecorded);
                        buffer.Add(chunk);
                    }
                };
                capture.RecordingStopped += (s, e) => tcs.TrySetResult(true);

                capture.StartRecording();
                await Task.Delay(durationMs);
                capture.StopRecording();

                // RecordingStopped が発火するまで最大2秒待つ
                await Task.WhenAny(tcs.Task, Task.Delay(2000));

                // バッファを結合
                var allBytes = buffer.SelectMany(b => b).ToArray();

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
            catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
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
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                    await InvokeAsyncMethod(preview, "SeekAsync", startFrame));
                await Task.Delay(300);

                // 2) 録音・再生・キャプチャを並行実行
                using var capture = new NAudio.Wave.WasapiLoopbackCapture();
                var waveFormat = capture.WaveFormat;
                var audioBuffer = new System.Collections.Concurrent.ConcurrentBag<byte[]>();
                var recordTcs = new TaskCompletionSource<bool>();
                capture.DataAvailable += (s, e) =>
                {
                    if (e.BytesRecorded > 0)
                    {
                        var chunk = new byte[e.BytesRecorded];
                        Buffer.BlockCopy(e.Buffer, 0, chunk, 0, e.BytesRecorded);
                        audioBuffer.Add(chunk);
                    }
                };
                capture.RecordingStopped += (s, e) => recordTcs.TrySetResult(true);
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
                await InvokeAsyncMethod(preview, "StopAsync");
                capture.StopRecording();
                await Task.WhenAny(recordTcs.Task, Task.Delay(2000));
                await captureTask;

                var allBytes = audioBuffer.SelectMany(b => b).ToArray();
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
            catch (Exception ex) { return (object)new { success = false, error = ex.Message }; }
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
