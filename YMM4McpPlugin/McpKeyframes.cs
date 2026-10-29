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
            (string itemId, int tf, int tl) selector;
            try
            {
                selector = KeyframeSelectorSafety.Parse(
                    req.QueryString["item_id"], req.QueryString["frame"], req.QueryString["layer"]);
            }
            catch (ArgumentException ex)
            {
                return Failure("ITEM_SELECTOR_INVALID", ex.Message);
            }
            var (itemId, tf, tl) = selector;
            string prop = req.QueryString["prop"] ?? "";
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var located = LocateTimelineItem(itemId, tf, tl);
                if (located.error != null) return located.error;
                var item = located.item!;
                var identity = GetItemIdentity(item);
                // The opt-in live smoke test must never certify a partial
                // inventory as complete when a YMM4 property getter fails.
                bool strictInventory = string.IsNullOrWhiteSpace(prop)
                    && req.QueryString["strict_inventory"] == "1";
                IEnumerable<string> properties;
                try
                {
                    properties = string.IsNullOrWhiteSpace(prop)
                        ? strictInventory ? ListAnimatablePropertiesStrict(item) : ListAnimatableProperties(item)
                        : new[] { prop };
                }
                catch (Exception)
                {
                    return Failure("ANIMATION_INVENTORY_INCOMPLETE",
                        "アニメーション一覧の取得に失敗しました。実機検証を中断してください");
                }
                var animations = new List<object>();
                var names = new List<string>();
                foreach (var name in properties)
                {
                    try
                    {
                        var anim = ResolveAnimation(item, name);
                        if (anim == null)
                        {
                            if (strictInventory)
                                return Failure("ANIMATION_INVENTORY_INCOMPLETE",
                                    "アニメーション " + name + " を再取得できませんでした");
                            continue;
                        }
                        names.Add(name);
                        animations.Add(new
                        {
                            prop = name,
                            type = anim.GetType().Name,
                            keyframes = ReadKeyframes(anim, located.frame)
                        });
                    }
                    catch (Exception)
                    {
                        if (strictInventory)
                            return Failure("ANIMATION_INVENTORY_INCOMPLETE",
                                "アニメーション " + name + " の読取に失敗しました");
                        throw;
                    }
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
                    inventory_complete = strictInventory,
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

                // Do not create a key beyond the item's visible duration.
                // The host can otherwise accept a frame which will not survive
                // a project save/reload or which targets another time span.
                if (action == "set" && !KeyframeMutationVerification.IsInItemRange(
                        relativeFrame, ReadItemInfo(item).length))
                    return Failure("KEYFRAME_OUT_OF_RANGE",
                        "キーフレーム位置がアイテムの長さの範囲外です。0以上、Length未満を指定してください");

                // KeyFrames belongs to the timeline item and is shared across
                // its animations.  Bind a detached Animation before adding a
                // nonzero keyframe; do not replace an existing binding.
                // Remove, clear and even a set at frame zero can mutate a
                // shared collection. Reject cross-item bindings for every action.
                if (KeyframeHostBinding.HasConflictingBinding(item, anim))
                    return Failure("KEYFRAME_BINDING_CONFLICT",
                        "Animationが別のアイテムのKeyFramesに接続されています。誤ったアイテムを変更しないため処理を中止しました");
                if (action == "set" && !KeyframeHostBinding.CanSetAt(item, anim, relativeFrame))
                    return Failure("KEYFRAME_BINDING_UNAVAILABLE",
                        "Animationをアイテムの共有KeyFramesに接続できませんでした。未接続のキーを追加しないため処理を中止しました");
                // Capture observable state before invoking reflection-based
                // native methods. They can partially mutate then return false.
                List<int> framesBefore;
                List<double> valuesBefore;
                try
                {
                    framesBefore = ReadFrameList(anim);
                    valuesBefore = ReadValueList(anim);
                }
                catch (Exception)
                {
                    // No native edit has started. Never mutate an animation
                    // whose initial state cannot be represented faithfully.
                    return Failure("KEYFRAME_READBACK_INVALID",
                        "変更前のAnimationに読取不能なFrames/Valuesがあります。編集せず中止しました");
                }
                bool changed;
                List<int> framesAfter;
                List<double> valuesAfter;
                double? defaultAfter;
                try
                {
                    if (action == "clear")
                        changed = ClearKeyframes(anim);
                    else if (action == "remove")
                        changed = RemoveKeyframe(anim, relativeFrame);
                    else
                        changed = UpsertKeyframe(anim, relativeFrame, value!.Value);
                    // Both readbacks must succeed before a native result is trusted.
                    framesAfter = ReadFrameList(anim);
                    valuesAfter = ReadValueList(anim);
                    // A throwing default getter must invalidate the revision
                    // too: it executes after the host has already mutated.
                    defaultAfter = action == "set" ? TryGetDefaultValue(anim) : null;
                }
                catch (Exception)
                {
                    // A native call may have modified the animation before
                    // throwing, or the host may refuse to expose its new state.
                    // Fail closed and invalidate any optimistic revision.
                    MarkItemChanged(item);
                    return Failure("KEYFRAME_PARTIAL_MUTATION",
                        "YMM4の操作または読戻しで例外が発生しました。部分変更の可能性があるため再試行前に状態を確認してください");
                }
                if (!changed)
                {
                    if (KeyframeMutationVerification.HasObservableChange(
                            framesBefore, valuesBefore, framesAfter, valuesAfter))
                    {
                        // Even a failed call can have changed the host model.
                        // Invalidate optimistic revisions and require inspection.
                        MarkItemChanged(item);
                        return Failure("KEYFRAME_PARTIAL_MUTATION",
                            "YMM4の操作が失敗しましたが、Animationが部分変更されています。再試行前に状態を確認してください");
                    }
                    return Failure("KEYFRAME_METHOD_UNAVAILABLE", "YMM4のAnimation/KeyFrames APIを呼べませんでした。inspectで署名を確認してください");
                }
                // A reflected native method can return successfully without
                // changing the requested frame/value. Do not acknowledge that
                // as a successful edit. Invalidate stale revisions because the
                // host may have been partially mutated despite the mismatch.
                if (!KeyframeMutationVerification.Confirm(
                        action, relativeFrame, value, framesAfter,
                        valuesAfter, defaultAfter) ||
                    !KeyframeMutationVerification.PreservesUnrelatedKeys(
                        action, relativeFrame, framesBefore, valuesBefore,
                        framesAfter, valuesAfter))
                {
                    MarkItemChanged(item);
                    return Failure("KEYFRAME_READBACK_MISMATCH",
                        "YMM4が操作成功を返しましたが、キーフレームの読戻しが一致しません。再試行前にinspectで状態を確認してください");
                }
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
            if (itemId.Length == 0 && (frame < 0 || layer < 0))
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
                // Do not stop at the first coordinate match: YMM4 can have
                // multiple overlapping items on the same frame/layer.
                var match = KeyframeSelectorSafety.FindUniqueCoordinateItem(
                    rawItems, frame, layer, iv =>
                    {
                        var info = ReadItemInfo(iv);
                        return (info.item, info.frame, info.layer);
                    });
                if (match.ambiguous)
                    return (null, 0, 0, Failure("ITEM_COORDINATES_AMBIGUOUS",
                        "同じframe/layerに複数のアイテムがあります。item_idで指定してください"));
                target = match.item;
                if (target != null)
                {
                    foundFrame = frame;
                    foundLayer = layer;
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

        // Strict mode is used only by the opt-in destructive smoke test.
        // Unlike GetPropObj, direct reflection does not hide throwing getters.
        // A failure in *any* readable item property aborts the test rather
        // than allowing an unobserved animation to be changed silently.
        private static string[] ListAnimatablePropertiesStrict(object item)
        {
            var names = new List<string>();
            foreach (var property in item.GetType().GetProperties(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                if (property.GetIndexParameters().Length != 0 || property.GetGetMethod(true) == null)
                    continue;
                object? value = property.GetValue(item);
                if (value != null && value.GetType().Name.Contains("ReactiveProperty"))
                {
                    var inner = value.GetType().GetProperty("Value",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (inner?.GetGetMethod(true) != null)
                        value = inner.GetValue(value);
                }
                if (IsAnimation(value)) names.Add(property.Name);
            }
            return names.Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
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
                result.Add(new { at, absolute_frame = KeyframeSelectorSafety.AbsoluteFrame(itemStartFrame, at), value });
            }
            return result;
        }

        private static List<int> ReadFrameList(object anim)
        {
            var keyFrames = GetPropObj(anim, "KeyFrames") ?? anim;
            var framesObj = GetPropObj(keyFrames, "Frames") ??
                GetPropObj(keyFrames, "FrameList") ?? GetPropEnum(keyFrames, "Frames");
            if (framesObj == null) return new List<int>();
            if (framesObj is not IEnumerable enumerable || framesObj is string)
                throw new InvalidOperationException("Animation Frames is not enumerable");
            return KeyframeMutationVerification.ParseObservedFrames(enumerable);
        }

        private static List<double> ReadValueList(object anim)
        {
            var valuesObj = GetPropObj(anim, "Values") ?? GetPropEnum(anim, "Values");
            if (valuesObj == null) return new List<double>();
            if (valuesObj is not IEnumerable enumerable || valuesObj is string)
                throw new InvalidOperationException("Animation Values is not enumerable");
            return KeyframeMutationVerification.ParseObservedValues(
                enumerable, entry => GetPropObj(entry, "Value") ?? entry);
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
            var before = ReadFrameList(anim);
            var valuesBefore = ReadValueList(anim);
            var valueCountBefore = valuesBefore.Count;
            bool isNew = frame > 0 && !before.Contains(frame);
            var keyFrames = GetPropObj(anim, "KeyFrames");

            // A host Add may insert a shared Frame without inserting its
            // AnimationValue. In that state Values.Count == Frames.Count can
            // look like a valid no-implicit-start layout, and TrySetValueAt
            // would overwrite the *original* start value and claim success.
            bool Complete() => !isNew || KeyframeInsertionSafety.IsComplete(
                before, valueCountBefore, ReadFrameList(anim), ReadValueList(anim).Count, frame);

            bool UndoIncomplete(bool force = false)
            {
                var currentFrames = ReadFrameList(anim);
                if (!isNew) return true;
                if (!currentFrames.Contains(frame))
                    return currentFrames.SequenceEqual(before)
                        && ReadValueList(anim).SequenceEqual(valuesBefore);
                if (!force && Complete()) return true;
                int insertedIndex = currentFrames.IndexOf(frame);
                int valueIndex = insertedIndex + valueCountBefore - before.Count;
                if (keyFrames != null)
                    TryInvokeNamed(keyFrames, new[] { "Remove", "RemoveKeyFrame", "RemoveKeyframe" }, frame);
                if (ReadFrameList(anim).Contains(frame))
                    TryInvokeNamed(anim, KeyframeRemoveNames, frame);
                // A host Remove may remove the frame but leave the extra
                // AnimationValue. Never retry with misaligned collections.
                if (ReadFrameList(anim).Contains(frame)) return false;
                if (ReadValueList(anim).Count != valueCountBefore)
                {
                    if (GetPropObj(anim, "Values") is not IList values
                        || !KeyframeInsertionSafety.TryRemoveInsertedValue(
                            values, valueCountBefore, valueIndex))
                        return false;
                }
                // A host may also mutate unrelated values or frames. Stop
                // instead of claiming that a failed edit was rolled back.
                return ReadFrameList(anim).SequenceEqual(before)
                    && ReadValueList(anim).SequenceEqual(valuesBefore);
            }

            if (TrySetValueAt(anim, frame, value) && Complete()) return true;
            if (TryInvokeNamed(anim, KeyframeMethodNames, frame, value) && Complete()) return true;
            if (!UndoIncomplete()) return false;

            if (!ReadFrameList(anim).Contains(frame))
                TryInvokeNamed(anim, KeyframeMethodNames, frame);
            if (Complete() && TrySetValueAt(anim, frame, value)) return true;
            if (!UndoIncomplete(force: true)) return false;

            if (keyFrames != null)
            {
                if (!ReadFrameList(anim).Contains(frame))
                    TryInvokeNamed(keyFrames, new[] { "Add", "AddKeyFrame", "Insert" }, frame);
                if (Complete() && TrySetValueAt(anim, frame, value)) return true;
                if (!UndoIncomplete(force: true)) return false;
            }
            // With the original layout restored, the mutable collection
            // fallback inserts the Frame and Value as one logical operation.
            return KeyframeCollectionMutation.TryUpsert(anim, frame, value);
        }

        private static bool RemoveKeyframe(object anim, int frame)
        {
            if (TryInvokeNamed(anim, KeyframeRemoveNames, frame)) return true;
            var keyFrames = GetPropObj(anim, "KeyFrames");
            if (keyFrames != null && TryInvokeNamed(keyFrames, new[] { "Remove", "RemoveKeyFrame" }, frame)) return true;
            return KeyframeCollectionRemoval.TryRemove(anim, frame);
        }

        private static bool ClearKeyframes(object anim)
        {
            // Prefer a single verified transaction over a series of removals.
            // Otherwise a failure halfway through leaves an edited animation.
            if (KeyframeCollectionRemoval.TryClear(anim)) return true;
            var frames = ReadFrameList(anim).OrderByDescending(f => f).ToArray();
            if (frames.Length == 0)
            {
                foreach (var name in new[] { "ClearKeyFrames", "Clear", "ResetKeyFrames" })
                    if (TryInvokeNamed(anim, new[] { name })) return true;
                return false;
            }
            bool any = false;
            foreach (var frame in frames)
            {
                if (!RemoveKeyframe(anim, frame))
                {
                    if (any)
                        throw new InvalidOperationException(
                            "Keyframe clear stopped after a partial native removal; inspect the animation before retrying.");
                    return false;
                }
                any = true;
            }
            if (ReadFrameList(anim).Count != 0)
                throw new InvalidOperationException(
                    "Native keyframe clear returned success but some frames remain.");
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
                {
                    int match = frames.FindIndex(f => f == frame);
                    index = frame == 0 && (frames.Count == 0 || frames[0] != 0)
                        ? 0 : match >= 0 ? match + 1 : -1;
                }
                else
                    index = frames.FindIndex(f => f == frame);
                if (index < 0 || index >= list.Count) return false;
                var entry = list[index];
                if (entry != null && KeyframeValueMutation.TrySet(entry, value))
                    return true;
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
            => KeyframeMethodInvoker.TryInvokeNamed(target, names, args);
    }
}
