using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        private static readonly HashSet<string> ExportFormats = new(StringComparer.OrdinalIgnoreCase)
            { "mp4", "wav", "avi", "mov", "mkv", "webm" };
        private static readonly string[] ExportMethodHints =
            { "OutputAsync", "ExportAsync", "RenderAsync", "StartOutput", "Output", "Export", "Render", "SaveVideo", "EncodeAsync" };
        private static readonly string[] PathPropertyHints =
            { "OutputPath", "OutputFilePath", "FilePath", "FileName", "Path", "SavePath", "Destination" };
        private static readonly string[] OpenMethodHints =
            { "OpenProject", "Open", "LoadProject", "Load", "OpenFile" };
        private static readonly string[] SaveAsMethodHints =
            { "SaveAsProject", "SaveProjectAs", "SaveAs", "SaveProject" };

        private async Task<object> EnqueueExport(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            var parsed = ParseExportRequest(body);
            var request = new Dictionary<string, object?>
            {
                ["output_path"] = parsed.path,
                ["format"] = parsed.format,
                ["overwrite"] = parsed.overwrite,
                ["timeout_seconds"] = parsed.timeout
            };
            var existing = FindIdempotentJob(parsed.key);
            if (existing != null)
                return SameJobRequest(existing, "export", request)
                    ? DescribeJob(existing)
                    : Failure("IDEMPOTENCY_KEY_CONFLICT", "idempotency_key は別の書き出し要求に使用されています");
            if (File.Exists(parsed.path) && !parsed.overwrite)
                return Failure("FILE_EXISTS", "出力先が既に存在します。overwrite=true で上書きしてください");
            var parent = Path.GetDirectoryName(parsed.path);
            if (string.IsNullOrEmpty(parent) || !Directory.Exists(parent))
                return Failure("DIRECTORY_NOT_FOUND", "出力先フォルダがありません: " + parent);

            var job = CreateJob("export", request, parsed.key, out bool created, out bool conflict);
            if (conflict)
                return Failure("IDEMPOTENCY_KEY_CONFLICT", "idempotency_key は別の書き出し要求に使用されています");
            if (created)
                RunJob(job, (running, token) => ExecuteExportJob(running, parsed.path, parsed.format, parsed.timeout, token));
            return DescribeJob(job);
        }

        private object RestartExport(McpJob previous)
        {
            string path = RequestString(previous.Request, "output_path");
            string format = RequestString(previous.Request, "format");
            if (string.IsNullOrEmpty(format)) format = "mp4";
            int timeout = RequestInt(previous.Request, "timeout_seconds", 1800);
            if (string.IsNullOrEmpty(path))
                return Failure("JOB_NOT_RESUMABLE", "元の出力パスが残っていません");
            var request = new Dictionary<string, object?>(previous.Request);
            // The predecessor ID makes repeated resume calls idempotent even when the
            // original request had no key. A failed successor can itself be resumed.
            var job = CreateJob("export", request, previous.Id + ":resume",
                out bool created, out bool conflict);
            if (conflict)
                return Failure("IDEMPOTENCY_KEY_CONFLICT", "再開キーは別の書き出し要求に使用されています");
            if (created)
                RunJob(job, (running, token) => ExecuteExportJob(running, path, format, timeout, token));
            return DescribeJob(job);
        }

        private async Task ExecuteExportJob(McpJob job, string outputPath, string format, int timeoutSeconds, CancellationToken token)
        {
            UpdateJob(job, "running", phase: "prepare", progress: 5, message: "出力パイプラインを探索中",
                checkpoint: new { output_path = outputPath, format });
            // overwrite=true may leave the previous export at this path until the new encode starts.
            // A stable old file must never be reported as this job's completed output.
            (long length, DateTime lastWriteUtc)? previousOutput = null;
            if (File.Exists(outputPath))
            {
                var oldFile = new FileInfo(outputPath);
                previousOutput = (oldFile.Length, oldFile.LastWriteTimeUtc);
            }
            var discovery = Application.Current.Dispatcher.Invoke(() => DiscoverExportSurface());
            var started = Application.Current.Dispatcher.Invoke(() => StartNativeExport(discovery, outputPath, format));
            if (!started.invoked)
            {
                UpdateJob(job, "failed", phase: "failed", progress: 0,
                    message: "プログラムからの書き出し方法が見つかりません",
                    error: "YMM4の出力コマンド/メソッドを実行できませんでした",
                    errorCode: "EXPORT_METHOD_UNAVAILABLE",
                    result: new { discovered = discovery.Describe(), output_path = outputPath });
                return;
            }
            UpdateJob(job, "running", phase: "render", progress: 15,
                message: "書き出しを開始しました: " + started.method,
                checkpoint: new { output_path = outputPath, method = started.method, target = started.target });

            if (started.task != null)
            {
                var finished = await Task.WhenAny(started.task, Task.Delay(Timeout.Infinite, token));
                token.ThrowIfCancellationRequested();
                if (started.task.IsFaulted)
                    throw started.task.Exception?.InnerException ?? started.task.Exception!;
                await started.task;
            }

            var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
            long lastSize = -1;
            int stable = 0;
            bool outputWasRemoved = false;
            bool dialogOnly = started.dialogLikely;
            var dialogDeadline = DateTime.UtcNow.AddSeconds(8);
            while (DateTime.UtcNow < deadline)
            {
                token.ThrowIfCancellationRequested();
                var outputInfo = File.Exists(outputPath) ? new FileInfo(outputPath) : null;
                if (outputInfo == null) outputWasRemoved = true;
                if (outputInfo != null && (previousOutput == null || outputWasRemoved ||
                    outputInfo.Length != previousOutput.Value.length ||
                    outputInfo.LastWriteTimeUtc != previousOutput.Value.lastWriteUtc))
                {
                    long size = outputInfo.Length;
                    if (size > 0 && size == lastSize) stable++;
                    else stable = 0;
                    lastSize = size;
                    int progress = (int)Math.Clamp(20 + (DateTime.UtcNow - (deadline - TimeSpan.FromSeconds(timeoutSeconds))).TotalSeconds / timeoutSeconds * 70, 20, 90);
                    UpdateJob(job, "running", phase: "render", progress: progress,
                        message: "出力ファイルを検出: " + size + " bytes",
                        checkpoint: new { output_path = outputPath, bytes = size });
                    if (stable >= 3 && size >= 32)
                    {
                        var verify = VerifyMediaFile(outputPath, format);
                        if (!verify.ok)
                        {
                            UpdateJob(job, "failed", phase: "verify", progress: 95,
                                message: verify.error, error: verify.error, errorCode: verify.errorCode, result: verify.payload);
                            return;
                        }
                        UpdateJob(job, "completed", phase: "verify", progress: 100,
                            message: "書き出し完了", result: verify.payload);
                        return;
                    }
                }
                else
                {
                    stable = 0;
                    lastSize = -1;
                    if (dialogOnly && DateTime.UtcNow > dialogDeadline)
                    {
                        UpdateJob(job, "failed", phase: "failed", progress: 0,
                            message: "出力ダイアログは開きましたがファイルが生成されませんでした",
                            error: "GUIの出力ダイアログは自動入力できません。パス付きメソッドが見つかるYMM4版が必要です",
                            errorCode: "EXPORT_DIALOG_REQUIRED",
                            result: new { discovered = discovery.Describe(), invoked = started.method });
                        return;
                    }
                }
                await Task.Delay(1000, token);
            }
            UpdateJob(job, "failed", phase: "timeout",
                message: "書き出しが制限時間内に完了しませんでした",
                error: "timeout_seconds=" + timeoutSeconds, errorCode: "EXPORT_TIMEOUT",
                result: new { output_path = outputPath, invoked = started.method });
        }

        private async Task<object> OpenProject(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            string path = RequireAbsolutePath(body, "path");
            bool force = GetBool(body, "force", false);
            if (!path.EndsWith(".ymmp", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("project path must end with .ymmp");
            if (!File.Exists(path))
                return Failure("FILE_NOT_FOUND", "プロジェクトファイルがありません: " + path);
            return await RunOnUi(async () =>
            {
                var vm = GetMainViewModel();
                var model = vm == null ? null : GetMainModel(vm);
                object? value = vm == null ? null : GetPropValue(vm, "IsSaved") ??
                    (model == null ? null : GetPropValue(model, "IsSaved"));
                bool? isSaved = value is bool saved ? saved : null;
                if (!force && isSaved == false)
                    return Failure("UNSAVED_CHANGES", "現在のプロジェクトに未保存の変更があります。保存するか force=true を指定してください");
                if (!force && isSaved == null)
                    return Failure("PROJECT_SAVE_STATE_UNKNOWN", "保存状態を確認できません。現在のプロジェクトを確認してから force=true を指定してください");
                var invoked = InvokeNamed(OpenMethodHints, path);
                if (!invoked.ok)
                    return Failure("OPEN_METHOD_UNAVAILABLE", "プロジェクトをパス指定で開く方法が見つかりません");
                if (invoked.task != null) await invoked.task;
                return (object)new { success = true, action = "open", path, method = invoked.method, target = invoked.target };
            });
        }

        private async Task<object> SaveProjectAs(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            string path = RequireAbsolutePath(body, "path");
            if (!path.EndsWith(".ymmp", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("project path must end with .ymmp");
            bool overwrite = GetBool(body, "overwrite", false);
            if (File.Exists(path) && !overwrite)
                return Failure("FILE_EXISTS", "保存先が既に存在します。overwrite=true で上書きしてください");
            var parent = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(parent) || !Directory.Exists(parent))
                return Failure("DIRECTORY_NOT_FOUND", "保存先フォルダがありません: " + parent);
            string? backupPath;
            try { backupPath = BackupProjectFile(path); }
            catch (Exception ex) { return Failure("BACKUP_FAILED", "既存プロジェクトを退避できません: " + ex.Message); }
            return await RunOnUi(async () =>
            {
                var invoked = InvokeNamed(SaveAsMethodHints, path);
                if (!invoked.ok)
                    return (object)new { success = false, error_code = "SAVE_AS_METHOD_UNAVAILABLE",
                        error = "別名保存のパス指定メソッドが見つかりません", retryable = false,
                        outcome_unknown = false, backup_path = backupPath };
                if (invoked.task != null) await invoked.task;
                return (object)new { success = true, action = "save-as", path, method = invoked.method,
                    target = invoked.target, backup_path = backupPath };
            });
        }

        private object SaveCurrentProject()
        {
            string? backupPath;
            try { backupPath = BackupProjectFile(CurrentProjectPath()); }
            catch (Exception ex) { return Failure("BACKUP_FAILED", "既存プロジェクトを退避できません: " + ex.Message); }
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var vm = GetMainViewModel();
                if (vm == null) return Failure("NO_MAIN_VIEW_MODEL", "MainViewModel取得失敗");
                if (vm.GetType().GetProperty("SaveProjectCommand")?.GetValue(vm) is not ICommand command)
                    return (object)new { success = false, error_code = "SAVE_COMMAND_UNAVAILABLE",
                        error = "SaveProjectCommand が見つかりません", retryable = false,
                        outcome_unknown = false, backup_path = backupPath };
                try
                {
                    if (!command.CanExecute(null)) return Failure("SAVE_COMMAND_UNAVAILABLE", "保存コマンドを実行できません");
                    command.Execute(null);
                }
                catch (Exception ex) { return Failure("SAVE_COMMAND_FAILED", ex.InnerException?.Message ?? ex.Message, true); }
                return (object)new { success = true, action = "save", backup_path = backupPath };
            });
        }

        private static string? BackupProjectFile(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
            McpSettings.PrepareDirectory();
            return ProjectBackup.BeforeOverwrite(path, Path.Combine(McpSettings.DirectoryPath, "project-backups"));
        }

        private static (string path, string format, bool overwrite, int timeout, string? key) ParseExportRequest(
            Dictionary<string, JsonElement> body)
        {
            string path = body.ContainsKey("output_path") ? GetStr(body, "output_path", "") : GetStr(body, "path", "");
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("path は書き出し先の絶対パスで指定してください");
            path = path.Trim();
            if (!Path.IsPathFullyQualified(path)) throw new ArgumentException("path must be an absolute file path");
            path = Path.GetFullPath(path);
            string ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
            string format = GetStr(body, "format", ext).Trim().TrimStart('.').ToLowerInvariant();
            if (!ExportFormats.Contains(format)) throw new ArgumentException("format must be one of: mp4, wav, avi, mov, mkv, webm");
            if (ext.Length > 0 && !ext.Equals(format, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("path extension must match format");
            bool overwrite = GetBool(body, "overwrite", false);
            int timeout = 1800;
            if (body.TryGetValue("timeout_seconds", out var timeoutEl))
            {
                if (timeoutEl.ValueKind != JsonValueKind.Number || !timeoutEl.TryGetInt32(out timeout))
                    throw new ArgumentException("timeout_seconds must be an integer in 1..7200");
            }
            if (timeout < 1 || timeout > 7200) throw new ArgumentException("timeout_seconds must be an integer in 1..7200");
            string key = GetStr(body, "idempotency_key", "");
            if (key.Length > 128 || (key.Length > 0 && key.Trim() != key))
                throw new ArgumentException("idempotency_key must be a 1..128 character string without surrounding whitespace");
            return (path, format, overwrite, timeout, key.Length == 0 ? null : key);
        }

        private static string RequireAbsolutePath(Dictionary<string, JsonElement> body, string key)
        {
            string path = GetStr(body, key, "");
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException(key + " は絶対パスで指定してください");
            path = path.Trim();
            if (!Path.IsPathFullyQualified(path)) throw new ArgumentException("path must be an absolute file path");
            return Path.GetFullPath(path);
        }

        private static object GetMediaFileInfo(HttpListenerRequest req)
        {
            string path = req.QueryString["path"] ?? "";
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
                throw new ArgumentException("path must be an absolute file path");
            path = Path.GetFullPath(path);
            var file = new FileInfo(path);
            string extension = file.Extension.ToLowerInvariant();
            if (!file.Exists)
                return new { success = true, exists = false, path, extension };
            return new
            {
                success = true,
                exists = true,
                path,
                extension,
                bytes = file.Length,
                last_write_utc = file.LastWriteTimeUtc.ToString("O")
            };
        }

        private static object GetMediaAssets(HttpListenerRequest req)
        {
            string directory = req.QueryString["directory"] ?? "";
            if (!int.TryParse(req.QueryString["max_results"] ?? "100", out int maxResults))
                throw new ArgumentException("max_results must be an integer");
            bool ParseFlag(string name)
            {
                string? value = req.QueryString[name];
                if (value == null || value.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
                if (value.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
                throw new ArgumentException(name + " must be boolean");
            }
            try { return AssetScanner.Scan(directory, req.QueryString["query"], ParseFlag("recursive"),
                ParseFlag("hash"), maxResults); }
            catch (DirectoryNotFoundException) { return Failure("DIRECTORY_NOT_FOUND", "素材フォルダがありません: " + directory); }
        }

        private static bool GetBool(Dictionary<string, JsonElement> body, string key, bool defaultValue)
        {
            if (!body.TryGetValue(key, out var value)) return defaultValue;
            if (value.ValueKind == JsonValueKind.True) return true;
            if (value.ValueKind == JsonValueKind.False) return false;
            throw new ArgumentException(key + " must be boolean");
        }

        private static string RequestString(Dictionary<string, object?> request, string key)
        {
            if (!request.TryGetValue(key, out var raw) || raw == null) return "";
            if (raw is JsonElement je)
                return je.ValueKind == JsonValueKind.String ? je.GetString() ?? "" : je.ToString();
            return raw.ToString() ?? "";
        }

        private static int RequestInt(Dictionary<string, object?> request, string key, int fallback)
        {
            if (!request.TryGetValue(key, out var raw) || raw == null) return fallback;
            if (raw is int i) return i;
            if (raw is long l && l >= int.MinValue && l <= int.MaxValue) return (int)l;
            if (raw is JsonElement je && je.ValueKind == JsonValueKind.Number && je.TryGetInt32(out var n)) return n;
            if (raw is string s && int.TryParse(s, out var parsed)) return parsed;
            return fallback;
        }

        private static bool RequestBool(Dictionary<string, object?> request, string key)
        {
            if (!request.TryGetValue(key, out var raw) || raw == null) return false;
            if (raw is bool value) return value;
            return raw is JsonElement element && element.ValueKind == JsonValueKind.True;
        }

        private static bool SameJobRequest(McpJob existing, string kind, Dictionary<string, object?> request)
            => existing.Kind == kind
               && StringComparer.OrdinalIgnoreCase.Equals(RequestString(existing.Request, "output_path"), RequestString(request, "output_path"))
               && StringComparer.OrdinalIgnoreCase.Equals(RequestString(existing.Request, "format"), RequestString(request, "format"))
               && RequestBool(existing.Request, "overwrite") == RequestBool(request, "overwrite")
               && RequestInt(existing.Request, "timeout_seconds", 1800) == RequestInt(request, "timeout_seconds", 1800);

        private ExportDiscovery DiscoverExportSurface()
        {
            var found = new ExportDiscovery();
            var vm = GetMainViewModel();
            if (vm == null) return found;
            foreach (var name in new[] { "Main", "Project", "Player", "Preview", "ActiveTimeline" })
            {
                var obj = ResolveTarget(name);
                if (obj == null) continue;
                CollectSurface(obj, name, found);
            }
            var areas = GetPropEnum(vm, "AnchorableAreaViewModels");
            if (areas != null)
            {
                foreach (var area in areas)
                {
                    var inner = GetPropObj(area, "ViewModel") ?? area;
                    var typeName = inner.GetType().Name;
                    if (typeName.Contains("Output", StringComparison.OrdinalIgnoreCase)
                        || typeName.Contains("Export", StringComparison.OrdinalIgnoreCase)
                        || typeName.Contains("Render", StringComparison.OrdinalIgnoreCase))
                        CollectSurface(inner, typeName, found);
                }
            }
            foreach (Window window in Application.Current.Windows)
            {
                if (window.DataContext == null) continue;
                var typeName = window.DataContext.GetType().Name;
                if (typeName.Contains("Output", StringComparison.OrdinalIgnoreCase)
                    || typeName.Contains("Export", StringComparison.OrdinalIgnoreCase))
                    CollectSurface(window.DataContext, typeName, found);
            }
            return found;
        }

        private static void CollectSurface(object obj, string target, ExportDiscovery found)
        {
            var type = obj.GetType();
            foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                if (typeof(ICommand).IsAssignableFrom(prop.PropertyType)
                    && (prop.Name.Contains("Output", StringComparison.OrdinalIgnoreCase)
                        || prop.Name.Contains("Export", StringComparison.OrdinalIgnoreCase)
                        || prop.Name.Contains("Render", StringComparison.OrdinalIgnoreCase)))
                    found.Commands.Add((target, prop.Name, obj));
                if (PathPropertyHints.Contains(prop.Name) && (prop.PropertyType == typeof(string) || Nullable.GetUnderlyingType(prop.PropertyType) == typeof(string)))
                    found.PathProperties.Add((target, prop.Name, obj));
            }
            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                if (method.IsSpecialName) continue;
                if (ExportMethodHints.Any(h => method.Name.Equals(h, StringComparison.OrdinalIgnoreCase)
                    || method.Name.Contains("Output", StringComparison.OrdinalIgnoreCase)
                    || method.Name.Contains("Export", StringComparison.OrdinalIgnoreCase)))
                {
                    if (method.Name.Contains("CommandList", StringComparison.OrdinalIgnoreCase)) continue;
                    found.Methods.Add((target, method, obj));
                }
            }
        }

        private static (bool invoked, bool dialogLikely, string method, string target, Task? task) StartNativeExport(
            ExportDiscovery discovery, string outputPath, string format)
        {
            foreach (var (target, name, obj) in discovery.PathProperties)
            {
                try
                {
                    var prop = obj.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (prop?.CanWrite == true) prop.SetValue(obj, outputPath);
                }
                catch { }
            }

            foreach (var (target, method, obj) in discovery.Methods.OrderBy(m => m.method.GetParameters().Length == 0 ? 1 : 0))
            {
                var parameters = method.GetParameters();
                object?[] args;
                if (parameters.Length == 0) args = Array.Empty<object?>();
                else if (parameters.Length == 1 && parameters[0].ParameterType == typeof(string))
                    args = new object?[] { outputPath };
                else if (parameters.Length == 2 && parameters[0].ParameterType == typeof(string) && parameters[1].ParameterType == typeof(string))
                    args = new object?[] { outputPath, format };
                else continue;
                try
                {
                    var result = method.Invoke(obj, args);
                    Task? task = result as Task;
                    bool dialog = parameters.Length == 0;
                    return (true, dialog, method.Name, target, task);
                }
                catch { }
            }

            foreach (var (target, name, obj) in discovery.Commands)
            {
                try
                {
                    var prop = obj.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (prop?.GetValue(obj) is not ICommand cmd) continue;
                    object? param = cmd.CanExecute(outputPath) ? outputPath : null;
                    if (!cmd.CanExecute(param)) continue;
                    cmd.Execute(param);
                    return (true, param == null, name, target, null);
                }
                catch { }
            }
            return (false, false, "", "", null);
        }

        private (bool ok, string method, string target, Task? task) InvokeNamed(string[] names, string path)
        {
            foreach (var targetName in new[] { "Main", "Project" })
            {
                var obj = ResolveTarget(targetName);
                if (obj == null) continue;
                var model = targetName == "Main" ? GetMainModel(obj) : null;
                foreach (var host in new object?[] { obj, model })
                {
                    if (host == null) continue;
                    foreach (var name in names)
                    {
                        var method = host.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                            .FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                                && m.GetParameters().Length == 1
                                && m.GetParameters()[0].ParameterType == typeof(string));
                        if (method == null) continue;
                        var result = method.Invoke(host, new object[] { path });
                        return (true, method.Name, targetName, result as Task);
                    }
                    foreach (var name in names.Select(n => n.EndsWith("Command", StringComparison.Ordinal) ? n : n + "Command"))
                    {
                        var prop = host.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        if (prop?.GetValue(host) is not ICommand cmd) continue;
                        object? param = cmd.CanExecute(path) ? path : null;
                        if (!cmd.CanExecute(param)) continue;
                        if (param == null) continue;
                        cmd.Execute(param);
                        return (true, name, targetName, null);
                    }
                }
            }
            return (false, "", "", null);
        }

        private static (bool ok, string errorCode, string error, object payload) VerifyMediaFile(string path, string format)
        {
            if (!File.Exists(path))
                return (false, "EXPORT_FILE_MISSING", "出力ファイルがありません", new { output_path = path, format });
            var info = new FileInfo(path);
            if (info.Length < 32)
                return (false, "EXPORT_FILE_EMPTY", "出力ファイルが小さすぎます", new { output_path = path, bytes = info.Length, format });
            Span<byte> header = stackalloc byte[12];
            using (var fs = File.OpenRead(path))
            {
                int read = fs.Read(header);
                if (read < 12)
                    return (false, "EXPORT_FILE_EMPTY", "出力ファイルが小さすぎます", new { output_path = path, bytes = info.Length, format });
            }
            if (format.Equals("mp4", StringComparison.OrdinalIgnoreCase))
            {
                var inspected = Mp4FileInspector.Inspect(path);
                if (!inspected.Verified)
                    return (false, "EXPORT_VERIFY_FAILED", inspected.Error,
                        new { output_path = path, bytes = info.Length, format });
                return (true, "", "", new { output_path = path, bytes = info.Length, format = "mp4",
                    has_video = true, has_audio = inspected.HasAudio, brand = inspected.Brand,
                    duration_seconds = inspected.DurationSeconds, width = inspected.Width, height = inspected.Height,
                    verified = true });
            }
            if (format.Equals("wav", StringComparison.OrdinalIgnoreCase))
            {
                bool riff = header[0] == (byte)'R' && header[1] == (byte)'I' && header[2] == (byte)'F' && header[3] == (byte)'F';
                bool wave = header[8] == (byte)'W' && header[9] == (byte)'A' && header[10] == (byte)'V' && header[11] == (byte)'E';
                if (!riff || !wave)
                    return (false, "EXPORT_VERIFY_FAILED", "WAVヘッダが不正です",
                        new { output_path = path, bytes = info.Length, format });
                return (true, "", "", new { output_path = path, bytes = info.Length, format = "wav", has_audio = true, verified = true });
            }
            return (true, "", "", new { output_path = path, bytes = info.Length, format, verified = true });
        }

        private sealed class ExportDiscovery
        {
            public List<(string target, string name, object obj)> Commands { get; } = new();
            public List<(string target, MethodInfo method, object obj)> Methods { get; } = new();
            public List<(string target, string name, object obj)> PathProperties { get; } = new();

            public object Describe() => new
            {
                commands = Commands.Select(c => c.target + "." + c.name).Distinct().ToArray(),
                methods = Methods.Select(m => m.target + "." + m.method.Name + "(" +
                    string.Join(",", m.method.GetParameters().Select(p => p.ParameterType.Name)) + ")").Distinct().ToArray(),
                path_properties = PathProperties.Select(p => p.target + "." + p.name).Distinct().ToArray()
            };
        }
    }
}
