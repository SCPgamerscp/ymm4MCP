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

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        private readonly SemaphoreSlim _editGate = new(1, 1);
        private static object Failure(string code, string message, bool outcomeUnknown = false)
            => new { success = false, error_code = code, error = message, retryable = false, outcome_unknown = outcomeUnknown };

        private static object[] TimelineObjects(object timeline)
        {
            var items = GetPropEnum(timeline, "Items")
                ?? throw new InvalidOperationException("Timeline items unavailable");
            return items.Cast<object>().Select(iv => GetPropObj(iv, "Item") ?? iv).ToArray();
        }

        private static object FindCharacter(object timeline, string name)
        {
            var characters = GetPropEnum(timeline, "Characters")?.Cast<object>().ToArray()
                ?? throw new InvalidOperationException("Character list unavailable");
            var matches = characters.Where(c => GetPropObj(c, "Name")?.ToString() == name).ToArray();
            if (matches.Length != 1)
                throw new ArgumentException("キャラ一覧から一意の完全一致名を指定してください: " + name);
            return matches[0];
        }

        private object GetCharacters() => Application.Current.Dispatcher.Invoke(() =>
        {
            var vm = GetMainViewModel();
            var timeline = vm == null ? null : GetPropObj(vm, "ActiveTimelineViewModel");
            if (timeline == null) return Failure("NO_TIMELINE", "プロジェクトを開いてください");
            var characters = GetPropEnum(timeline, "Characters");
            if (characters == null) return Failure("CHARACTERS_UNAVAILABLE", "キャラ一覧を取得できません");
            return (object)new { success = true, characters = characters.Cast<object>().Select(c => new
            {
                name = GetPropObj(c, "Name")?.ToString(),
                layer = GetPropObj(c, "Layer"),
                voice_plugin = GetPropObj(GetPropObj(c, "Voice") ?? c, "API")?.ToString(),
                voice_name = GetPropObj(GetPropObj(c, "Voice") ?? c, "Display")?.ToString(),
                style = GetPropObj(GetPropObj(c, "VoiceParameter") ?? c, "Style")?.ToString(),
                note = "Voice settings are inherited from the registered character. Missing metadata is null."
            }).ToArray() };
        });

        private static int NonNegative(Dictionary<string, JsonElement> body, string key, int defaultValue = 0)
        {
            if (!body.TryGetValue(key, out var value)) return defaultValue;
            if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out int result) || result < 0)
                throw new ArgumentException(key + " must be a nonnegative 32-bit integer");
            return result;
        }

        private static bool MatchesItemType(object item, string typeName)
        {
            var name = item.GetType().Name;
            return name.Equals(typeName, StringComparison.Ordinal)
                || name.EndsWith(typeName, StringComparison.OrdinalIgnoreCase)
                || name.Contains(typeName, StringComparison.OrdinalIgnoreCase);
        }

        private static MethodInfo? FindAddMethod(object model, string name, bool voice, bool character)
        {
            var methods = model.GetType()
                .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(m => m.Name == name)
                .ToArray();
            if (methods.Length == 0) return null;

            int want = voice ? 5 : 3;
            MethodInfo? Unique(Func<MethodInfo, bool> pred)
            {
                var hits = methods.Where(pred).ToArray();
                return hits.Length == 1 ? hits[0] : hits.FirstOrDefault(m => m.IsPublic) ?? (hits.Length > 0 ? hits[0] : null);
            }

            return Unique(m =>
            {
                var p = m.GetParameters();
                return p.Length == want && p[0].ParameterType == typeof(int) && p[1].ParameterType == typeof(int)
                    && (character || p[2].ParameterType == typeof(string));
            }) ?? Unique(m =>
            {
                var p = m.GetParameters();
                return p.Length >= 3 && p[0].ParameterType == typeof(int) && p[1].ParameterType == typeof(int);
            }) ?? (methods.Length == 1 ? methods[0] : methods.FirstOrDefault(m => m.IsPublic) ?? methods[0]);
        }

        private static object? EmptyDecorations(ParameterInfo parameter)
        {
            var type = parameter.ParameterType;
            Type? element = type.IsArray ? type.GetElementType()
                : type.IsGenericType ? type.GetGenericArguments().FirstOrDefault()
                : null;
            return element != null ? Array.CreateInstance(element, 0) : null;
        }

        private static object?[] BuildAddParameters(MethodInfo method, int frame, int layer, object third, string? text, bool voice)
        {
            var parameters = method.GetParameters();
            var args = new object?[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
            {
                var t = parameters[i].ParameterType;
                if (i == 0 && t == typeof(int)) args[i] = frame;
                else if (i == 1 && t == typeof(int)) args[i] = layer;
                else if (t == typeof(string)) args[i] = text ?? third as string ?? "";
                else if (voice && i >= 4) args[i] = EmptyDecorations(parameters[i]);
                else if (i == 2) args[i] = third;
                else if (i == 3 && text != null) args[i] = text;
                else if (parameters[i].HasDefaultValue) args[i] = parameters[i].DefaultValue;
                else args[i] = t.IsValueType ? Activator.CreateInstance(t) : null;
            }
            return args;
        }

        private static bool TryAddViaViewModel(object vm, string kind, string typeName, int frame, int layer, string value, object? character)
        {
            string method = kind == "face" ? "AddFaceItem" : "Add" + typeName;
            object payload = character ?? value;
            if (TryMethod(vm, method, frame, layer, payload)) return true;
            if (TryMethod(vm, method, payload, frame, layer)) return true;
            if (TryMethod(vm, method, value, frame, layer)) return true;
            return TryCmd(vm, method + "Command") || TryCmd(vm, "Add" + typeName + "Command");
        }

        private static void TryRecordHistory(object model)
        {
            try
            {
                var history = GetPropObj(model, "UndoRedoManager");
                history?.GetType().GetMethod("Record", Type.EmptyTypes)?.Invoke(history, null);
            }
            catch { }
        }

        private object DescribeAddedItem(object item, string? warning, bool verified, int requestedFrame, int requestedLayer, string? character)
        {
            var info = ReadItemInfo(item);
            long end = (long)info.frame + Math.Max(info.length, 0);
            if (end > int.MaxValue) warning = JoinWarning(warning, "endFrame が Int32 範囲を超えています");
            if (info.length <= 0) warning = JoinWarning(warning, "追加アイテムの実長が未確定です。再追加せず items で確認してください");
            return new
            {
                success = true,
                verified,
                warning,
                type = info.type,
                frame = info.frame,
                layer = info.layer,
                length = info.length,
                endFrame = end <= int.MaxValue ? (int)end : info.frame,
                text = info.text,
                character,
                requestedFrame,
                requestedLayer
            };
        }

        private static string? JoinWarning(string? current, string extra)
            => string.IsNullOrEmpty(current) ? extra : current + " / " + extra;

        // Compare object identity before/after. Post-add helpers (history, length, type name)
        // must not turn a successful edit into success=false — that causes AI retries and duplicates.
        private async Task<object> AddNativeItem(HttpListenerRequest request, string kind)
        {
            var body = await ReadBody(request);
            var supported = new HashSet<string> { "frame", "layer", "length", "path", "text", "character" };
            if (body.Keys.Any(k => !supported.Contains(k))) throw new ArgumentException("Unsupported item parameter");
            int frame = NonNegative(body, "frame"), layer = NonNegative(body, "layer");
            int? length = body.ContainsKey("length") ? NonNegative(body, "length") : null;
            if (length == 0 || (length.HasValue && (long)frame + length > int.MaxValue))
                throw new ArgumentException("length must be positive and endFrame must fit in Int32");
            bool voice = kind == "voice", characterItem = voice || kind == "tachie" || kind == "face";
            bool media = kind == "video" || kind == "audio" || kind == "image";
            string value = GetStr(body, media ? "path" : "text", "");
            string character = GetStr(body, "character", "");
            if (voice && length.HasValue) throw new ArgumentException("Voice length is determined by synthesis");
            if (kind == "image" && !length.HasValue) throw new ArgumentException("image requires length");
            if ((kind == "text" || voice) && string.IsNullOrWhiteSpace(value)) throw new ArgumentException("text is required");
            if (media)
            {
                if (!Path.IsPathFullyQualified(value)) throw new ArgumentException("path must be absolute");
                value = Path.GetFullPath(value);
                if (!File.Exists(value)) return Failure("FILE_NOT_FOUND", "素材ファイルがありません: " + value);
            }
            if (characterItem && string.IsNullOrWhiteSpace(character)) throw new ArgumentException("character is required");

            string typeName = kind switch
            {
                "voice" => "VoiceItem",
                "text" => "TextItem",
                "video" => "VideoItem",
                "audio" => "AudioItem",
                "image" => "ImageItem",
                "tachie" => "TachieItem",
                "face" => "TachieFaceItem",
                _ => throw new ArgumentException("Unknown item kind")
            };

            object? vm = null, timeline = null, model = null;
            HashSet<object>? before = null;
            bool invoked = false;
            try
            {
                object? pending = Application.Current.Dispatcher.Invoke(() =>
                {
                    vm = GetMainViewModel() ?? throw new InvalidOperationException("MainViewModel unavailable");
                    timeline = GetPropObj(vm, "ActiveTimelineViewModel") ?? throw new InvalidOperationException("No timeline");
                    model = GetMainModel(vm) ?? throw new InvalidOperationException("MainModel unavailable");
                    before = new HashSet<object>(TimelineObjects(timeline), ReferenceEqualityComparer.Instance);
                    object third = characterItem ? FindCharacter(timeline, character) : value;
                    var methodNames = kind == "face"
                        ? new[] { "AddFaceItem", "AddTachieFaceItem" }
                        : new[] { "Add" + typeName + (voice ? "Async" : ""), "Add" + typeName, "Add" + typeName + "Async" };
                    MethodInfo? method = null;
                    foreach (var name in methodNames)
                    {
                        method = FindAddMethod(model, name, voice, characterItem);
                        if (method != null) break;
                    }
                    invoked = true;
                    if (method != null)
                        return method.Invoke(model, BuildAddParameters(method, frame, layer, third, voice || kind == "text" ? value : null, voice));
                    if (!TryAddViaViewModel(vm, kind, typeName, frame, layer, value, characterItem ? third : null))
                        throw new NotSupportedException("Add method signature unavailable for " + typeName);
                    return null;
                });
                if (pending is Task task) await task;
                return Application.Current.Dispatcher.Invoke(() => FinishAdd(vm!, timeline!, model!, before!, typeName, length, frame, layer, characterItem ? character : null));
            }
            catch (Exception ex)
            {
                if (invoked)
                {
                    try
                    {
                        return Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (vm == null || timeline == null || before == null)
                                return Failure("ADD_OUTCOME_UNKNOWN",
                                    "追加処理の途中で失敗しました。再試行せず items で確認してください: " + (ex.InnerException?.Message ?? ex.Message), true);
                            return FinishAdd(vm, timeline, model ?? vm, before, typeName, length, frame, layer, characterItem ? character : null, ex);
                        });
                    }
                    catch (Exception inspectEx)
                    {
                        return Failure("ADD_OUTCOME_UNKNOWN",
                            "追加処理の途中で失敗しました。再試行せず items で確認してください: " + (inspectEx.InnerException?.Message ?? inspectEx.Message), true);
                    }
                }
                return Failure(ex is ArgumentException ? "INVALID_ARGUMENT" : "ADD_FAILED",
                    ex.InnerException?.Message ?? ex.Message);
            }
        }

        private object FinishAdd(object vm, object timeline, object model, HashSet<object> before, string typeName,
            int? length, int requestedFrame, int requestedLayer, string? character, Exception? postAddError = null)
        {
            if (!ReferenceEquals(GetPropObj(vm, "ActiveTimelineViewModel"), timeline))
            {
                var moved = TryDescribeNewcomers(timeline, before, typeName, length, requestedFrame, requestedLayer, character, model,
                    JoinWarning("処理中にタイムラインが変更されました", postAddError?.InnerException?.Message ?? postAddError?.Message));
                if (moved != null) return moved;
                return Failure("TIMELINE_CHANGED", "処理中にタイムラインが変更されました。再試行せず結果を確認してください", true);
            }

            var described = TryDescribeNewcomers(timeline, before, typeName, length, requestedFrame, requestedLayer, character, model,
                postAddError == null ? null : "追加後の後処理で例外: " + (postAddError.InnerException?.Message ?? postAddError.Message));
            if (described != null) return described;
            if (postAddError != null)
                return Failure("ADD_OUTCOME_UNKNOWN",
                    "追加処理は実行されましたが結果を確認できません。再試行せず items で確認してください: " + (postAddError.InnerException?.Message ?? postAddError.Message), true);
            return Failure("ADD_NOT_VERIFIED", "追加メソッドは実行されましたが新しいアイテムが見つかりません。再試行せず items で確認してください", true);
        }

        private object? TryDescribeNewcomers(object timeline, HashSet<object> before, string typeName, int? length,
            int requestedFrame, int requestedLayer, string? character, object model, string? warning)
        {
            object[] newcomers;
            try { newcomers = TimelineObjects(timeline).Where(i => !before.Contains(i)).ToArray(); }
            catch { return null; }
            if (newcomers.Length == 0) return null;

            var typed = newcomers.Where(i => MatchesItemType(i, typeName)).ToArray();
            object item;
            bool verified;
            if (typed.Length == 1)
            {
                item = typed[0];
                verified = true;
            }
            else if (newcomers.Length == 1)
            {
                item = newcomers[0];
                verified = false;
                warning = JoinWarning(warning, "型名が想定と異なるため検証を緩和しました: " + item.GetType().Name);
            }
            else if (typed.Length > 1)
            {
                item = typed[0];
                verified = false;
                warning = JoinWarning(warning, $"新規アイテムが{typed.Length}件あります。再追加せず items で確認してください");
            }
            else
            {
                item = newcomers[0];
                verified = false;
                warning = JoinWarning(warning, $"新規アイテムが{newcomers.Length}件あります。再追加せず items で確認してください");
            }

            if (length.HasValue)
            {
                try
                {
                    var lengthProperty = item.GetType().GetProperty("Length");
                    if (lengthProperty?.CanWrite != true)
                        warning = JoinWarning(warning, "追加されましたが長さを設定できません");
                    else
                        lengthProperty.SetValue(item, length.Value);
                }
                catch (Exception ex)
                {
                    warning = JoinWarning(warning, "追加されましたが長さの設定に失敗: " + (ex.InnerException?.Message ?? ex.Message));
                }
            }

            TryRecordHistory(model);
            return DescribeAddedItem(item, warning, verified, requestedFrame, requestedLayer, character);
        }
    }
}
