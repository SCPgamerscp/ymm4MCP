using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace YMM4McpPlugin
{
    // Validate that a newly inserted frame has a corresponding new value.
    // Count parity alone is insufficient: a frame-only Add can change an
    // implicit-start layout (0 frames, 1 value) into an apparently valid
    // no-implicit-start layout (1 frame, 1 value).
    internal static class KeyframeInsertionSafety
    {
        // Some YMM4 host methods insert both a frame and a default value.
        // If the subsequent value write fails, the extra value must not be
        // left behind after the shared frame has been removed.
        internal static bool TryRemoveInsertedValue(IList values, int previousCount, int index)
        {
            try
            {
                if (values.Count == previousCount) return true;
                if (values.Count != previousCount + 1 || values.IsReadOnly
                    || values.IsFixedSize || index < 0 || index >= values.Count)
                    return false;
                values.RemoveAt(index);
                return values.Count == previousCount;
            }
            catch
            {
                // Host IList implementations can throw even from Count or
                // IsReadOnly. Treat that as an unsafe rollback, not success.
                return false;
            }
        }

        private static bool StrictlyIncreasing(IReadOnlyList<int> frames)
        {
            int previous = -1;
            for (int i = 0; i < frames.Count; i++)
            {
                int current = frames[i];
                if (current <= previous) return false;
                previous = current;
            }
            return true;
        }

        internal static bool IsComplete(
            IReadOnlyList<int> beforeFrames, int beforeValues,
            IReadOnlyList<int> afterFrames, int afterValues, int insertedFrame)
        {
            try
            {
                if (insertedFrame <= 0 || beforeFrames.Contains(insertedFrame))
                    return false;
                // A malformed or reordered source cannot be certified by
                // comparing counts and the apparent insertion position.
                if (!StrictlyIncreasing(beforeFrames) || !StrictlyIncreasing(afterFrames))
                    return false;
                int offset = beforeValues - beforeFrames.Count;
                if (offset != 0 && offset != 1) return false;
                if (afterFrames.Count != beforeFrames.Count + 1
                    || afterValues != beforeValues + 1
                    || afterValues - afterFrames.Count != offset)
                    return false;
                int insertionIndex = 0;
                while (insertionIndex < beforeFrames.Count
                    && beforeFrames[insertionIndex] < insertedFrame)
                    insertionIndex++;
                if (afterFrames[insertionIndex] != insertedFrame) return false;
                for (int i = 0; i < beforeFrames.Count; i++)
                    if (afterFrames[i < insertionIndex ? i : i + 1] != beforeFrames[i])
                        return false;
                return true;
            }
            catch
            {
                // Reflection-backed host lists may throw while being read.
                // Never treat a partially observed insertion as committed.
                return false;
            }
        }
    }
}
