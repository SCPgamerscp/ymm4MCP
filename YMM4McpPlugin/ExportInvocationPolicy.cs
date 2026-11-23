using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Input;

namespace YMM4McpPlugin
{
    internal static class ExportInvocationPolicy
    {
        private static readonly string[] FormatNames = { "Exo", "Mp4", "Wav", "Avi", "Mov", "Mkv", "Webm" };

        internal static bool TryCaptureOutputBaseline(string path, bool overwrite,
            out (long length, DateTime lastWriteUtc)? previousOutput)
        {
            previousOutput = null;
            if (Directory.Exists(path)) return false;
            var file = new FileInfo(path);
            if (!file.Exists) return true;
            if (!overwrite) return false;
            previousOutput = (file.Length, file.LastWriteTimeUtc);
            return true;
        }

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

        internal static bool ReturnsCompletionTask(MethodInfo method)
        {
            Type type = method.ReturnType;
            return !method.ContainsGenericParameters &&
                (typeof(Task).IsAssignableFrom(type) || type == typeof(ValueTask) || IsGenericValueTask(type));
        }

        private static bool IsGenericValueTask(Type type) =>
            type.IsGenericType && type.GetGenericTypeDefinition() == typeof(ValueTask<>);

        internal static Task? GetCompletionTask(object? result)
        {
            if (result is Task task) return task;
            // Convert a ValueTask exactly once. Some awaitable sources are single-use.
            try
            {
                if (result is ValueTask valueTask) return valueTask.AsTask();
                if (result != null && IsGenericValueTask(result.GetType()))
                    return (Task?)result.GetType().GetMethod(nameof(ValueTask.AsTask), Type.EmptyTypes)
                        ?.Invoke(result, null);
            }
            catch (Exception ex) when (ex is InvalidOperationException or TargetInvocationException)
            {
                // The call may have started the encoder even though its completion cannot be observed.
            }
            return null;
        }

        internal static MethodInfo? FindAwaitableCommand(ICommand command)
        {
            // The standard ICommand.Execute returns void. An explicit toolkit
            // interface or public ExecuteAsync can supply the actual Task.
            var type = command.GetType();
            var contract = type.GetInterface("CommunityToolkit.Mvvm.Input.IAsyncRelayCommand");
            var method = contract?.GetMethod("ExecuteAsync", new[] { typeof(object) });
            if (method != null && ReturnsCompletionTask(method)) return method;
            return type.GetMethods(BindingFlags.Public | BindingFlags.Instance).FirstOrDefault(candidate =>
                candidate.Name == "ExecuteAsync" && ReturnsCompletionTask(candidate) &&
                candidate.GetParameters() is [{ ParameterType: var parameter }] &&
                (parameter == typeof(string) || parameter == typeof(object)));
        }

        internal static Task? InvokeProjectPathMethod(MethodInfo method, object host, string path)
        {
            var task = GetCompletionTask(method.Invoke(host, new object[] { path }));
            if (task == null && ReturnsCompletionTask(method))
                throw new InvalidOperationException("YMM4 project operation did not return an observable completion task.");
            return task;
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
