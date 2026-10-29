using System;
using System.Globalization;

namespace YMM4McpPlugin
{
    // Keep GET /api/items/keyframes selectors as strict as mutation selectors.
    // In particular, int.TryParse(out ...) must never turn an omitted or
    // malformed coordinate into zero and silently select a different item.
    internal static class KeyframeSelectorSafety
    {
        internal static (string itemId, int frame, int layer) Parse(
            string? itemId, string? frame, string? layer)
        {
            string id = itemId ?? "";
            if (id.Length > 0)
            {
                if (string.IsNullOrWhiteSpace(id) || id.Length > 256)
                    throw new ArgumentException("item_id must be a non-empty identifier of at most 256 characters");
                foreach (char c in id)
                    if (char.IsControl(c) || c is '\u202a' or '\u202b' or '\u202c'
                        or '\u202d' or '\u202e' or '\u2066' or '\u2067' or '\u2068' or '\u2069')
                        throw new ArgumentException("item_id contains unsafe control characters");
                // item_id is authoritative; optional coordinate hints were
                // historically ignored and are kept backward compatible.
                return (id, -1, -1);
            }
            if (frame == null || layer == null)
                throw new ArgumentException("item_id or both frame and layer are required");
            return ("", ParseCoordinate(frame, "frame"), ParseCoordinate(layer, "layer"));
        }

        private static int ParseCoordinate(string value, string name)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int result))
                throw new ArgumentException(name + " must be a non-negative 32-bit integer");
            return result;
        }

        internal static bool IsAmbiguousMatch(object? selected, object? candidate)
            => selected != null && candidate != null && !ReferenceEquals(selected, candidate);

        internal static long AbsoluteFrame(int itemStartFrame, int relativeFrame)
            => (long)itemStartFrame + relativeFrame;
    }
}
