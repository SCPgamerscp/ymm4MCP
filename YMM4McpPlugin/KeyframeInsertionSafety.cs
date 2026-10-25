using System;
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
        internal static bool IsComplete(
            IReadOnlyList<int> beforeFrames, int beforeValues,
            IReadOnlyList<int> afterFrames, int afterValues, int insertedFrame)
        {
            if (insertedFrame <= 0 || beforeFrames.Contains(insertedFrame))
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
    }
}
