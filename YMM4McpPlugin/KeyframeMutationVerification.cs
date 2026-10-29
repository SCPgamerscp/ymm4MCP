using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    // Independent of YMM4's concrete Animation types so native reflection
    // paths and collection fallbacks share the same readback contract.
    internal static class KeyframeMutationVerification
    {
        // Do not silently omit malformed host entries during keyframe readback.
        // Dropping one Frame or Value can turn a partial native mutation into a
        // seemingly valid collection and cause the API to report false success.
        internal static List<int> ParseObservedFrames(System.Collections.IEnumerable source)
        {
            var frames = new List<int>();
            foreach (object? entry in source)
            {
                if (entry == null ||
                    !int.TryParse(Convert.ToString(entry, System.Globalization.CultureInfo.InvariantCulture),
                        System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out int frame) ||
                    frame < 0)
                    throw new InvalidOperationException("Animation contains an unreadable frame");
                frames.Add(frame);
            }
            return frames;
        }

        internal static List<double> ParseObservedValues(
            System.Collections.IEnumerable source, Func<object, object?> unwrap)
        {
            var values = new List<double>();
            foreach (object? entry in source)
            {
                if (entry == null)
                    throw new InvalidOperationException("Animation contains a null value");
                object? raw = unwrap(entry);
                if (raw == null ||
                    !double.TryParse(Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double value) ||
                    !double.IsFinite(value))
                    throw new InvalidOperationException("Animation contains an unreadable value");
                values.Add(value);
            }
            return values;
        }

        internal static bool IsInItemRange(int relativeFrame, int length)
            => length > 0 && relativeFrame >= 0 && relativeFrame < length;

        internal static bool HasObservableChange(
            IReadOnlyList<int> beforeFrames, IReadOnlyList<double> beforeValues,
            IReadOnlyList<int> afterFrames, IReadOnlyList<double> afterValues)
        {
            if (beforeFrames.Count != afterFrames.Count ||
                beforeValues.Count != afterValues.Count) return true;
            for (int i = 0; i < beforeFrames.Count; i++)
                if (beforeFrames[i] != afterFrames[i]) return true;
            for (int i = 0; i < beforeValues.Count; i++)
                if (!beforeValues[i].Equals(afterValues[i])) return true;
            return false;
        }

        // A readable inventory must describe a complete, unambiguous set of
        // frame/value pairs. A malformed layout must not be presented as
        // valid animation data to callers or to the strict live smoke test.
        internal static bool HasReadableLayout(
            IReadOnlyList<int> frames, IReadOnlyList<double> values)
            => frames != null && values != null &&
               TrySnapshot(frames, values, out _);

        internal static bool Confirm(string action, int at, double? requested,
            IReadOnlyList<int> frames, IReadOnlyList<double> values, double? defaultValue)
        {
            if (frames == null || values == null || at < 0) return false;
            var unique = new HashSet<int>();
            foreach (var frame in frames)
                if (frame < 0 || !unique.Add(frame)) return false;

            // A native remove can drop a frame while leaving its value behind.
            // Never acknowledge a removal with a misaligned collection.
            bool aligned = values.Count == frames.Count || values.Count == frames.Count + 1;
            if (action == "clear") return frames.Count == 0 && values.Count <= 1;
            if (action == "remove") return !unique.Contains(at) && aligned;
            if (action != "set" || !requested.HasValue ||
                !double.IsFinite(requested.Value)) return false;

            // Some host animations expose only a DefaultValue/Value property
            // and no Values array. Frame zero is still verifiable there.
            if (frames.Count == 0 && values.Count == 0)
                return at == 0 && SameValue(defaultValue, requested.Value);

            bool implicitStart = values.Count == frames.Count + 1;
            if (values.Count != frames.Count && !implicitStart)
                return false;

            int index = -1;
            if (at == 0 && implicitStart && !unique.Contains(0))
                index = 0;
            else
            {
                for (int i = 0; i < frames.Count; i++)
                    if (frames[i] == at)
                    {
                        index = i + (implicitStart ? 1 : 0);
                        break;
                    }
            }
            if (index < 0 || index >= values.Count) return false;
            return SameValue(values[index], requested.Value);
        }

        // A successful native call can update the requested key *and* silently
        // overwrite/delete another key. Compare logical frame/value pairs,
        // rather than raw list positions (which shift on insert/remove).
        internal static bool PreservesUnrelatedKeys(
            string action, int at,
            IReadOnlyList<int> beforeFrames, IReadOnlyList<double> beforeValues,
            IReadOnlyList<int> afterFrames, IReadOnlyList<double> afterValues)
        {
            if (action == "clear") return true; // Clearing all keys is intentional.
            if (action != "set" && action != "remove") return false;
            if (!TrySnapshot(beforeFrames, beforeValues, out var before) ||
                !TrySnapshot(afterFrames, afterValues, out var after))
                return false;
            foreach (var pair in before)
            {
                if (pair.Key == at) continue;
                if (!after.TryGetValue(pair.Key, out var value) ||
                    !pair.Value.Equals(value))
                    return false;
            }
            // A native method must not introduce unrelated keyframes either.
            foreach (var frame in after.Keys)
                if (frame != at && !before.ContainsKey(frame))
                    return false;
            return true;
        }

        private static bool TrySnapshot(
            IReadOnlyList<int> frames, IReadOnlyList<double> values,
            out Dictionary<int, double> snapshot)
        {
            snapshot = new Dictionary<int, double>();
            bool implicitStart = values.Count == frames.Count + 1;
            if (values.Count != frames.Count && !implicitStart) return false;
            if (implicitStart)
            {
                if (!double.IsFinite(values[0])) return false;
                snapshot[0] = values[0];
            }
            for (int i = 0; i < frames.Count; i++)
            {
                int frame = frames[i];
                double value = values[i + (implicitStart ? 1 : 0)];
                if (frame < 0 || !double.IsFinite(value) ||
                    !snapshot.TryAdd(frame, value))
                    return false;
            }
            return true;
        }

        private static bool SameValue(double? observed, double requested)
        {
            if (!observed.HasValue || !double.IsFinite(observed.Value)) return false;
            // YMM4 may store a float even when the API accepts a double.
            // Allow float precision while rejecting a different key value.
            return Math.Abs(observed.Value - requested) <=
                Math.Max(0.000001, Math.Abs(requested) * 0.000001);
        }
    }
}
