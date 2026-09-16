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

        private static MethodInfo RequireAddMethod(object model, string name, bool voice, bool character)
        {
            var methods = model.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.Name == name).Where(m =>
                {
                    var p = m.GetParameters();
                    return p.Length == (voice ? 5 : 3) && p[0].ParameterType == typeof(int) &&
                        p[1].ParameterType == typeof(int) && (character || p[2].ParameterType == typeof(string));
                }).ToArray();
            return methods.Length == 1 ? methods[0] : throw new NotSupportedException(name + " signature unavailable");
        }

        // Uses the known MainModel methods rather than falling back to unrelated UI commands.
        // Compare object identity before/after to avoid mistaking an existing item for the new one.
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
            bool invoked = false;
            try
            {
                string typeName = kind switch { "voice" => "VoiceItem", "text" => "TextItem", "video" => "VideoItem", "audio" => "AudioItem", "image" => "ImageItem", "tachie" => "TachieItem", "face" => "TachieFaceItem", _ => throw new ArgumentException("Unknown item kind") };
                var setup = Application.Current.Dispatcher.Invoke(() =>
                {
                    var vm = GetMainViewModel() ?? throw new InvalidOperationException("MainViewModel unavailable");
                    var timeline = GetPropObj(vm, "ActiveTimelineViewModel") ?? throw new InvalidOperationException("No timeline");
                    var model = GetMainModel(vm) ?? throw new InvalidOperationException("MainModel unavailable");
                    string methodName = kind == "face" ? "AddFaceItem" : "Add" + typeName + (voice ? "Async" : "");
                    var method = RequireAddMethod(model, methodName, voice, characterItem);
                    var before = new HashSet<object>(TimelineObjects(timeline), ReferenceEqualityComparer.Instance);
                    object third = characterItem ? FindCharacter(timeline, character) : value;
                    object?[] parameters;
                    if (voice)
                    {
                        var decorationType = method.GetParameters()[4].ParameterType.GetGenericArguments().Single();
                        parameters = new object?[] { frame, layer, third, value, Array.CreateInstance(decorationType, 0) };
                    }
                    else parameters = new object?[] { frame, layer, third };
                    invoked = true;
                    object? pending = method.Invoke(model, parameters);
                    return (vm, timeline, model, before, pending);
                });
                if (setup.pending is Task task) await task;
                return Application.Current.Dispatcher.Invoke(() =>
                {
                    if (!ReferenceEquals(GetPropObj(setup.vm, "ActiveTimelineViewModel"), setup.timeline))
                        return Failure("TIMELINE_CHANGED", "処理中にタイムラインが変更されました。結果を確認してください", true);
                    var added = TimelineObjects(setup.timeline).Where(i => !setup.before.Contains(i) && i.GetType().Name == typeName).ToArray();
                    if (added.Length != 1) return Failure("ADD_NOT_VERIFIED", "追加結果を一意に特定できません。itemsで確認してください", true);
                    var item = added[0];
                    if (length.HasValue)
                    {
                        var lengthProperty = item.GetType().GetProperty("Length");
                        if (lengthProperty?.CanWrite != true) return Failure("LENGTH_UNSUPPORTED", "追加されましたが長さを設定できません", true);
                        lengthProperty.SetValue(item, length.Value);
                    }
                    // Property setters participate in the host's edit history; record the boundary.
                    var history = GetPropObj(setup.model, "UndoRedoManager");
                    history?.GetType().GetMethod("Record", Type.EmptyTypes)?.Invoke(history, null);
                    var info = ReadItemInfo(item);
                    if (info.length <= 0 || (long)info.frame + info.length > int.MaxValue)
                        return Failure("INVALID_ADDED_RANGE", "追加アイテムの実長を確認できません", true);
                    return (object)new { success = true, verified = true, type = info.type, frame = info.frame,
                        layer = info.layer, length = info.length, endFrame = info.frame + info.length,
                        text = info.text, character = characterItem ? character : null,
                        requestedFrame = frame, requestedLayer = layer };
                });
            }
            catch (Exception ex)
            {
                return Failure(ex is ArgumentException ? "INVALID_ARGUMENT" : "ADD_FAILED",
                    ex.InnerException?.Message ?? ex.Message, invoked);
            }
        }
    }
}
