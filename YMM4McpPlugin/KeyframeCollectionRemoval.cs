using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace YMM4McpPlugin
{
    // Atomic fallback for YMM4 Animation layouts that expose Frames/Values
    // but do not provide a working native RemoveKeyFrame/ClearKeyFrames method.
    // Never silently report success when only one of the two lists changed.
    internal static class KeyframeCollectionRemoval
    {
        internal static bool TryRemove(object animation, int frame) =>
            Mutate(animation, frame, clear: false);

        internal static bool TryClear(object animation) =>
            Mutate(animation, 0, clear: true);

        private static bool Mutate(object animation, int frame, bool clear)
        {
            if (frame < 0) return false;
            var owner = Read(animation, "KeyFrames") ?? animation;
            if ((Read(owner, "Frames") ?? Read(owner, "FrameList")) is not IList frames
                || Read(animation, "Values") is not IList values)
                return false;

            object?[] beforeFrames;
            object?[] beforeValues;
            try
            {
                beforeFrames = frames.Cast<object?>().ToArray();
                beforeValues = values.Cast<object?>().ToArray();
            }
            catch { return false; }

            int offset = beforeValues.Length == beforeFrames.Length + 1 ? 1
                : beforeValues.Length == beforeFrames.Length && beforeFrames.Length > 0 ? 0 : -1;
            if (offset < 0) return false;
            var positions = new List<int>();
            int previous = -1;
            for (int i = 0; i < beforeFrames.Length; i++)
            {
                if (!int.TryParse(Convert.ToString(beforeFrames[i], CultureInfo.InvariantCulture),
                        NumberStyles.Integer, CultureInfo.InvariantCulture, out int current)
                    || current <= previous || current < 0 || (offset == 1 && current == 0))
                    return false;
                previous = current;
                if (clear || current == frame) positions.Add(i);
            }
            if (!clear && (positions.Count != 1 || (offset == 1 && frame == 0)))
                return false;
            // An implicit initial value must survive clear. A layout with no
            // implicit value cannot be reduced to an empty Values collection.
            if (offset == 0 && (clear || positions.Count == beforeFrames.Length))
                return false;
            if (positions.Count == 0) return clear && offset == 1;

            var removed = new HashSet<int>(positions);
            var afterFrames = beforeFrames.Where((_, i) => !removed.Contains(i)).ToArray();
            var afterValues = beforeValues.Where((_, i) => i < offset || !removed.Contains(i - offset)).ToArray();

            bool mutable = !frames.IsReadOnly && !frames.IsFixedSize
                && !values.IsReadOnly && !values.IsFixedSize;
            bool immutable = (frames.IsReadOnly || frames.IsFixedSize)
                && (values.IsReadOnly || values.IsFixedSize);
            if (!mutable && !immutable) return false;
            if (mutable)
            {
                bool threw = false;
                try
                {
                    foreach (int index in positions.OrderByDescending(i => i))
                    {
                        values.RemoveAt(index + offset);
                        frames.RemoveAt(index);
                    }
                }
                catch { threw = true; }
                if (!threw && Matches(frames, afterFrames) && Matches(values, afterValues))
                    return true;

                bool framesRestored = RestoreDeletions(frames, beforeFrames);
                bool valuesRestored = RestoreDeletions(values, beforeValues);
                if (!framesRestored || !valuesRestored)
                    throw new InvalidOperationException(
                        "Keyframe removal failed and the original collections could not be restored.");
                return false;
            }

            var frameProperty = owner.GetType().GetProperty("Frames", BindingFlags.Public | BindingFlags.Instance)
                ?? owner.GetType().GetProperty("FrameList", BindingFlags.Public | BindingFlags.Instance);
            var valueProperty = animation.GetType().GetProperty("Values", BindingFlags.Public | BindingFlags.Instance);
            if (frameProperty?.SetMethod?.IsPublic != true
                || valueProperty?.SetMethod?.IsPublic != true)
                return false;
            try
            {
                if (!ReferenceEquals(frameProperty.GetValue(owner), frames)
                    || !ReferenceEquals(valueProperty.GetValue(animation), values))
                    return false;
            }
            catch { return false; }

            object nextFrames = frames, nextValues = values;
            foreach (int index in positions.OrderByDescending(i => i))
            {
                if (!ImmutableRemoveAt(nextFrames, index, out nextFrames)
                    || !ImmutableRemoveAt(nextValues, index + offset, out nextValues))
                    return false;
            }
            if (nextFrames is not IList newFrames || nextValues is not IList newValues
                || !frameProperty.PropertyType.IsInstanceOfType(nextFrames)
                || !valueProperty.PropertyType.IsInstanceOfType(nextValues)
                || !Matches(newFrames, afterFrames) || !Matches(newValues, afterValues))
                return false;

            try
            {
                valueProperty.SetValue(animation, nextValues);
                frameProperty.SetValue(owner, nextFrames);
                if (ReferenceEquals(valueProperty.GetValue(animation), nextValues)
                    && ReferenceEquals(frameProperty.GetValue(owner), nextFrames))
                    return true;
            }
            catch { /* Restore both original references below. */ }

            try { frameProperty.SetValue(owner, frames); } catch { }
            try { valueProperty.SetValue(animation, values); } catch { }
            bool restoredFrames = false, restoredValues = false;
            try { restoredFrames = ReferenceEquals(frameProperty.GetValue(owner), frames); } catch { }
            try { restoredValues = ReferenceEquals(valueProperty.GetValue(animation), values); } catch { }
            if (!restoredFrames || !restoredValues)
                throw new InvalidOperationException(
                    "Immutable keyframe removal failed and original collections could not be restored.");
            return false;
        }

        private static bool ImmutableRemoveAt(object source, int index, out object updated)
        {
            updated = source;
            try
            {
                var method = source.GetType().GetMethod("RemoveAt", new[] { typeof(int) });
                if (method == null) return false;
                var result = method.Invoke(source, new object[] { index });
                if (result is not IList || ReferenceEquals(result, source)) return false;
                updated = result;
                return true;
            }
            catch { return false; }
        }

        private static bool Matches(IList list, object?[] expected)
        {
            try
            {
                if (list.Count != expected.Length) return false;
                for (int i = 0; i < expected.Length; i++)
                    if (!Equals(list[i], expected[i])) return false;
                return true;
            }
            catch { return false; }
        }

        // Only restore a verified subsequence (i.e. a series of deletions).
        // Do not overwrite unrelated entries if a non-conforming IList
        // changed the wrong values rather than removing elements.
        private static bool RestoreDeletions(IList list, object?[] before)
        {
            if (Matches(list, before)) return true;
            try
            {
                if (list.Count > before.Length) return false;
                int matched = 0;
                foreach (var original in before)
                {
                    if (matched < list.Count && Equals(list[matched], original))
                        matched++;
                }
                if (matched != list.Count) return false;
                for (int i = 0; i < before.Length; i++)
                {
                    if (i < list.Count && Equals(list[i], before[i])) continue;
                    try { list.Insert(i, before[i]); }
                    catch { /* An Insert can commit and then throw. */ }
                    if (i >= list.Count || !Equals(list[i], before[i]))
                        return false;
                }
                return Matches(list, before);
            }
            catch { return false; }
        }

        private static object? Read(object source, string property)
        {
            try
            {
                return source.GetType().GetProperty(property,
                    BindingFlags.Public | BindingFlags.Instance)?.GetValue(source);
            }
            catch { return null; }
        }
    }
}
