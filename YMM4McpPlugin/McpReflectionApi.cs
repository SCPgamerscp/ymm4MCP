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

    }
}
