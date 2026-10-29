using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    // Independent of YMM4's concrete Animation types so native reflection
    // paths and collection fallbacks share the same readback contract.
    internal static class KeyframeMutationVerification
    {
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
