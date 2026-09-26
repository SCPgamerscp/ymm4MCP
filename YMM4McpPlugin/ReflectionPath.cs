using System.Collections;
using System.Reflection;

namespace YMM4McpPlugin;

internal static class ReflectionPath
{
    internal static bool TryResolve(object? source, string path, out object? value)
    {
        value = source;
        if (source == null) return false;
        if (path.Length == 0) return true;
        const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
        foreach (string segment in path.Split('.'))
        {
            if (value == null || segment.Length == 0) return false;
            string member = segment;
            int? index = null;
            int bracket = segment.IndexOf('[');
            if (bracket >= 0)
            {
                if (!segment.EndsWith(']') || !int.TryParse(segment.AsSpan(bracket + 1, segment.Length - bracket - 2), out int parsed) || parsed < 0)
                    return false;
                member = segment[..bracket];
                index = parsed;
            }
            if (member.Length == 0 && index == null) return false;
            if (member.Length > 0)
            {
                try
                {
                    var type = value.GetType();
                    var property = type.GetProperty(member, flags);
                    var field = property == null ? type.GetField(member, flags) : null;
                    if (property == null && field == null) return false;
                    value = property != null ? property.GetValue(value) : field!.GetValue(value);
                    if (value != null && value.GetType().Name.Contains("ReactiveProperty"))
                    {
                        var inner = value.GetType().GetProperty("Value");
                        if (inner != null) value = inner.GetValue(value);
                    }
                }
                catch { return false; }
            }
            if (index != null)
            {
                if (value is not IEnumerable entries || value is string) return false;
                bool found = false;
                int position = 0;
                foreach (var entry in entries)
                {
                    if (position++ != index.Value) continue;
                    value = entry;
                    found = true;
                    break;
                }
                if (!found) return false;
            }
        }
        return true;
    }
}
