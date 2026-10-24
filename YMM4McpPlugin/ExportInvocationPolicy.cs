using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace YMM4McpPlugin
{
    internal static class ExportInvocationPolicy
    {
        private static readonly string[] FormatNames = { "Exo", "Mp4", "Wav", "Avi", "Mov", "Mkv", "Webm" };

        internal static bool MatchesFormat(string name, string requestedFormat)
        {
            foreach (var formatName in FormatNames)
            {
                // "Movie" is a generic video name, not the MOV container.
                if (formatName == "Mov" && name.Contains("Movie", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (name.Contains(formatName, StringComparison.OrdinalIgnoreCase)
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
                && (parameters.Length == 1 || parameters[1].ParameterType == typeof(string));
        }

        internal static bool HasDialogOnlyMethod(IEnumerable<MethodInfo> methods, string requestedFormat) =>
            methods.Any(method => MatchesFormat(method.Name, requestedFormat)
                && method.GetParameters().Length == 0);
    }
}
