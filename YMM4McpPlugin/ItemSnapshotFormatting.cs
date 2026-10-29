using System;
using System.Globalization;

namespace YMM4McpPlugin
{
    /// <summary>Keep item revision fingerprints stable across UI/OS cultures.</summary>
    internal static class ItemSnapshotFormatting
    {
        internal static string Format(object? value)
        {
            if (value == null) return "";
            if (value is IFormattable formattable)
                return formattable.ToString(null, CultureInfo.InvariantCulture) ?? "";
            return value.ToString() ?? "";
        }
    }
}
