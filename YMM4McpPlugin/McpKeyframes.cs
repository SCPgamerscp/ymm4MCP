using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace YMM4McpPlugin
{
    public partial class McpHttpServer
    {
        private static readonly string[] KeyframeMethodNames = { "AddKeyFrame", "AddKeyframe", "SetKeyFrame", "ChangeKeyFrame" };
        private static readonly string[] KeyframeRemoveNames = { "RemoveKeyFrame", "RemoveKeyframe" };

        private object GetItemKeyframes(HttpListenerRequest req)
        {
            int tf = -1, tl = -1;
            int.TryParse(req.QueryString["frame"], out tf);
            int.TryParse(req.QueryString["layer"], out tl);
            string itemId = req.QueryString["item_id"] ?? "";
            string prop = req.QueryString["prop"] ?? "";
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var located = LocateTimelineItem(itemId, tf, tl);
                if (located.error != null) return located.error;
                var item = located.item!;
                var identity = GetItemIdentity(item);
                var properties = string.IsNullOrWhiteSpace(prop)
                    ? ListAnimatableProperties(item)
                    : new[] { prop };
                var animations = new List<object>();
                var names = new List<string>();
                foreach (var name in properties)
                {
                    var anim = ResolveAnimation(item, name);
                    if (anim == null) continue;
                    names.Add(name);
                    animations.Add(new
                    {
                        prop = name,
                        type = anim.GetType().Name,
                        keyframes = ReadKeyframes(anim, located.frame)
                    });
                }
                return (object)new
                {
                    success = true,
                    item_id = identity.id,
                    revision = GetItemRevision(item),
                    identity_persistent = identity.persistent,
                    frame = located.frame,
                    layer = located.layer,
                    properties = names,
                    animations
                };
            });
        }

        private async Task<object> SetItemKeyframe(HttpListenerRequest req)
        {
            var body = await ReadBody(req);
            int tf = GetInt(body, "frame", -1);
            int tl = GetInt(body, "layer", -1);
            string itemId = GetStr(body, "item_id", "");
            string expectedRevision = GetStr(body, "expected_revision", "");
            if (expectedRevision.Length > 0 && itemId.Length == 0)
                return Failure("REVISION_REQUIRES_ITEM_ID", "expected_revision を使う場合は item_id も指定してください");
            TimelineInputValidation.RequireItemTarget(body);
            string prop = GetStr(body, "prop", "");
            if (string.IsNullOrWhiteSpace(prop)) throw new ArgumentException("prop is required");
            string action = GetStr(body, "action", "set").ToLowerInvariant();
            if (action is not ("set" or "remove" or "clear"))
                throw new ArgumentException("action must be set, remove, or clear");
            int relativeFrame = body.ContainsKey("at") ? NonNegative(body, "at")
                : body.ContainsKey("keyframe") ? NonNegative(body, "keyframe") : 0;
            double? value = action == "set" ? GetDouble(body, "value", double.NaN) : null;
            if (action == "set" && (value is null || !double.IsFinite(value.Value)))
                throw new ArgumentException("value must be a finite number");
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var located = LocateTimelineItem(itemId, tf, tl);
                if (located.error != null) return located.error;
                var item = located.item!;
                var conflict = RevisionConflict(item, expectedRevision);
                if (conflict != null) return conflict;
                var anim = ResolveAnimation(item, prop);
                if (anim == null)
                    return Failure("KEYFRAME_UNSUPPORTED", "プロパティ '" + prop + "' はAnimationではありません");

                bool changed;
                if (action == "clear")
                    changed = ClearKeyframes(anim);
                else if (action == "remove")
                    changed = RemoveKeyframe(anim, relativeFrame);
                else
                    changed = UpsertKeyframe(anim, relativeFrame, value!.Value);
                if (!changed)
                    return Failure("KEYFRAME_METHOD_UNAVAILABLE", "YMM4のAnimation/KeyFrames APIを呼べませんでした。inspectで署名を確認してください");
                MarkItemChanged(item);
                var vm = GetMainViewModel();
                if (vm != null)
                {
                    var model = GetMainModel(vm);
                    if (model != null) TryRecordHistory(model);
                }
                var identity = GetItemIdentity(item);
                return (object)new
                {
                    success = true,
                    item_id = identity.id,
                    revision = GetItemRevision(item),
                    identity_persistent = identity.persistent,
                    prop,
                    action,
                    at = relativeFrame,
                    keyframes = ReadKeyframes(anim, located.frame)
                };
            });
        }

        private (object? item, int frame, int layer, object? error) LocateTimelineItem(string itemId, int frame, int layer)
        {
            if (itemId.Length == 0 && frame < 0 && layer < 0)
                return (null, 0, 0, Failure("ITEM_SELECTOR_REQUIRED", "item_id または frame+layer を指定してください"));
            var vm = GetMainViewModel();
            if (vm == null) return (null, 0, 0, Failure("NO_MAIN_VIEW_MODEL", "VM失敗"));
            var tvm = GetPropObj(vm, "ActiveTimelineViewModel");
            if (tvm == null) return (null, 0, 0, Failure("NO_TIMELINE", "TVM失敗"));
            var rawItems = GetPropEnum(tvm, "Items");
            if (rawItems == null) return (null, 0, 0, Failure("ITEMS_UNAVAILABLE", "Items失敗"));
            object? target = null;
            int foundFrame = 0, foundLayer = 0;
            if (itemId.Length > 0)
            {
                target = FindItemById(rawItems, itemId, out bool ambiguous);
                if (ambiguous) return (null, 0, 0, Failure("ITEM_ID_AMBIGUOUS", "item_id が複数のアイテムに一致しました"));
            }
            else
            {
                foreach (var iv in rawItems)
                {
                    var info = ReadItemInfo(iv);
                    if ((frame < 0 || info.frame == frame) && (layer < 0 || info.layer == layer))
                    {
                        target = info.item;
                        foundFrame = info.frame;
                        foundLayer = info.layer;
                        break;
                    }
                }
            }
            if (target == null)
                return (null, 0, 0, Failure("ITEM_NOT_FOUND", itemId.Length > 0 ? "item_id に一致するアイテムがありません" : $"アイテム未発見 frame={frame} layer={layer}"));
            if (itemId.Length > 0)
            {
                var info = ReadItemInfo(target);
                foundFrame = info.frame;
                foundLayer = info.layer;
            }
            return (target, foundFrame, foundLayer, null);
        }

        private static IEnumerable<string> ListAnimatableProperties(object item)
        {
            return item.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(p => p.GetIndexParameters().Length == 0 && IsAnimation(GetPropObj(item, p.Name)))
                .Select(p => p.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToArray();
        }

        private static object? ResolveAnimation(object item, string prop)
        {
            var value = GetPropObj(item, prop);
            if (value != null && value.GetType().Name.Contains("ReactiveProperty"))
                value = GetPropObj(value, "Value") ?? value;
            return IsAnimation(value) ? value : null;
        }

        private static bool IsAnimation(object? value)
        {
            if (value == null) return false;
            var t = value.GetType();
            if (t.Name == "Animation" || t.Name.Contains("Animation")) return true;
            return GetPropObj(value, "KeyFrames") != null || t.GetMethod("GetValue") != null;
        }

        private static List<object> ReadKeyframes(object anim, int itemStartFrame)
        {
            var frames = ReadFrameList(anim);
            var values = ReadValueList(anim);
            var result = new List<object>();
            bool valuesHasStart = values.Count == frames.Count + 1;
            if (frames.Count == 0)
            {
                result.Add(new { at = 0, absolute_frame = itemStartFrame, value = values.Count > 0 ? values[0] : TryGetDefaultValue(anim) });
                return result;
            }
            if (valuesHasStart)
                result.Add(new { at = 0, absolute_frame = itemStartFrame, value = (double?)values[0] });
            for (int i = 0; i < frames.Count; i++)
            {
                int at = frames[i];
                if (valuesHasStart && at == 0) continue;
                int valueIndex = valuesHasStart ? i + 1 : i;
                double? value = valueIndex >= 0 && valueIndex < values.Count ? values[valueIndex] : null;
                result.Add(new { at, absolute_frame = itemStartFrame + at, value });
            }
            return result;
        }

        private static List<int> ReadFrameList(object anim)
        {
            var keyFrames = GetPropObj(anim, "KeyFrames") ?? anim;
            var framesObj = GetPropObj(keyFrames, "Frames") ?? GetPropObj(keyFrames, "FrameList") ?? GetPropEnum(keyFrames, "Frames");
            var list = new List<int>();
            if (framesObj is IEnumerable enumerable && framesObj is not string)
            {
                foreach (var x in enumerable)
                {
                    if (x == null) continue;
                    if (int.TryParse(Convert.ToString(x, CultureInfo.InvariantCulture), out int frame) && frame >= 0)
                        list.Add(frame);
                }
            }
            return list;
        }

        private static List<double> ReadValueList(object anim)
        {
            var valuesObj = GetPropObj(anim, "Values") ?? GetPropEnum(anim, "Values");
            var list = new List<double>();
            if (valuesObj is IEnumerable enumerable && valuesObj is not string)
            {
                foreach (var x in enumerable)
                {
                    if (x == null) continue;
                    var inner = GetPropObj(x, "Value") ?? x;
                    if (double.TryParse(Convert.ToString(inner, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                        && double.IsFinite(value))
                        list.Add(value);
                }
            }
            return list;
        }

        private static double? TryGetDefaultValue(object anim)
        {
            foreach (var name in new[] { "DefaultValue", "Value", "CurrentValue" })
            {
                var v = GetPropObj(anim, name);
                if (v != null && double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out double d)
                    && double.IsFinite(d))
                    return d;
            }
            return null;
        }

        private static bool UpsertKeyframe(object anim, int frame, double value)
        {
            if (TryInvokeNamed(anim, KeyframeMethodNames, frame, value)) return true;
            if (TryInvokeNamed(anim, KeyframeMethodNames, frame) && TrySetValueAt(anim, frame, value)) return true;
            var keyFrames = GetPropObj(anim, "KeyFrames");
            if (keyFrames != null)
            {
                if (TryInvokeNamed(keyFrames, new[] { "Add", "AddKeyFrame", "Insert" }, frame) && TrySetValueAt(anim, frame, value))
                    return true;
            }
            return TrySetValueAt(anim, frame, value);
        }

        private static bool RemoveKeyframe(object anim, int frame)
        {
            if (TryInvokeNamed(anim, KeyframeRemoveNames, frame)) return true;
            var keyFrames = GetPropObj(anim, "KeyFrames");
            return keyFrames != null && TryInvokeNamed(keyFrames, new[] { "Remove", "RemoveKeyFrame" }, frame);
        }

        private static bool ClearKeyframes(object anim)
        {
            var frames = ReadFrameList(anim).OrderByDescending(f => f).ToArray();
            if (frames.Length == 0)
            {
                foreach (var name in new[] { "ClearKeyFrames", "Clear", "ResetKeyFrames" })
                    if (TryInvokeNamed(anim, new[] { name })) return true;
                return false;
            }
            bool any = false;
            foreach (var frame in frames)
                any |= RemoveKeyframe(anim, frame);
            return any;
        }

        private static bool TrySetValueAt(object anim, int frame, double value)
        {
            var values = GetPropObj(anim, "Values");
            if (values is IList list && list.Count > 0)
            {
                var frames = ReadFrameList(anim);
                int index;
                if (frames.Count + 1 == list.Count)
                    index = frame == 0 && (frames.Count == 0 || frames[0] != 0) ? 0 : frames.FindIndex(f => f == frame) + 1;
                else
                    index = frames.FindIndex(f => f == frame);
                if (index < 0 || index >= list.Count) return false;
                var entry = list[index];
                if (entry != null)
                {
                    var vp = entry.GetType().GetProperty("Value");
                    if (vp?.CanWrite == true)
                    {
                        vp.SetValue(entry, Convert.ChangeType(value, vp.PropertyType, CultureInfo.InvariantCulture));
                        return true;
                    }
                }
            }
            foreach (var name in new[] { "DefaultValue", "Value" })
            {
                var p = anim.GetType().GetProperty(name);
                if (p?.CanWrite == true && frame == 0)
                {
                    p.SetValue(anim, Convert.ChangeType(value, p.PropertyType, CultureInfo.InvariantCulture));
                    return true;
                }
            }
            return false;
        }

        private static bool TryInvokeNamed(object target, IEnumerable<string> names, params object[] args)
        {
            var methods = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (var name in names)
            {
                foreach (var method in methods.Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    var parameters = method.GetParameters();
                    if (parameters.Length < args.Length) continue;
                    try
                    {
                        var call = new object?[parameters.Length];
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            if (i < args.Length)
                                call[i] = Convert.ChangeType(args[i], parameters[i].ParameterType, CultureInfo.InvariantCulture);
                            else if (parameters[i].HasDefaultValue)
                                call[i] = parameters[i].DefaultValue;
                            else
                                call = null;
                        }
                        if (call == null) continue;
                        method.Invoke(target, call);
                        return true;
                    }
                    catch { }
                }
            }
            return false;
        }
    }
}
