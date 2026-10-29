using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace YMM4McpPlugin
{
    internal static class ExportInvocationPolicy
    {
        private static readonly string[] FormatNames = { "Exo", "Mp4", "Wav", "Avi", "Mov", "Mkv", "Webm" };

        // Format names must be whole CamelCase/acronym tokens. Substrings such as
        // "avi" in "Saving" or "wav" in "Waving" are not export formats.
        private static bool HasFormatToken(string name, string format)
        {
            for (int i = 0; i <= name.Length - format.Length; i++)
            {
                if (!name.AsSpan(i, format.Length).Equals(format.AsSpan(), StringComparison.OrdinalIgnoreCase))
                    continue;
                if (i > 0 && char.IsLetterOrDigit(name[i - 1]) && !char.IsUpper(name[i]))
                    continue;
                int end = i + format.Length;
                if (end < name.Length && char.IsLower(name[end]))
                    continue;
                return true;
            }
            return false;
        }

        internal static bool MatchesFormat(string name, string requestedFormat)
        {
            // "Movie" is a generic video exporter, not a MOV exporter and not
            // an audio-only WAV exporter.
            if (requestedFormat.Equals("wav", StringComparison.OrdinalIgnoreCase)
                && name.Contains("Movie", StringComparison.OrdinalIgnoreCase)
                && !HasFormatToken(name, "Wav"))
                return false;
            foreach (var formatName in FormatNames)
            {
                if (formatName == "Mov" && name.Contains("Movie", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (HasFormatToken(name, formatName)
                    && !formatName.Equals(requestedFormat, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        internal static bool AcceptsPath(MethodInfo method, string requestedFormat)
        {
            if (!MatchesFormat(method.Name, requestedFormat)) return false;
            var parameters = method.GetParameters();
            return (parameters.Length == 1 || parameters.Length == 2)
                && parameters[0].ParameterType == typeof(string)
                // A second string argument is only safe when it is a format
                // selector. Passing "mp4" as a preset/output-directory path is unsafe.
                && (parameters.Length == 1 ||
                    (parameters[1].ParameterType == typeof(string)
                     && new[] { "format", "fileFormat", "outputFormat", "containerFormat", "extension" }
                         .Contains(parameters[1].Name, StringComparer.OrdinalIgnoreCase)));
        }

        // Match the MCP server's idempotency-key contract.  Internal whitespace is
        // allowed, but ASCII control characters must never reach persisted job keys.
        internal static bool IsValidIdempotencyKey(string key) =>
            key.Length <= 128
            && (key.Length == 0 || key.Trim() == key)
            && !key.Any(c => c <= '\u001f' || c == '\u007f');

        internal static bool HasDialogOnlyMethod(IEnumerable<MethodInfo> methods, string requestedFormat) =>
            methods.Any(method => MatchesFormat(method.Name, requestedFormat)
                && method.GetParameters().Length == 0);
    }
}
