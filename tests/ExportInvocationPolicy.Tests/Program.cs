using System.Reflection;
using YMM4McpPlugin;

static MethodInfo Method(string name, params Type[] parameters) =>
    typeof(ExportMethods).GetMethod(name, parameters) ?? throw new Exception($"Missing method: {name}");

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

Check(!ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputExoFile)), "mp4"),
    "MP4 must not invoke the EXO dialog");
Check(!ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputMp4File)), "mp4"),
    "A parameterless MP4 method cannot honor the requested path");
Check(ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputMp4File), typeof(string)), "mp4"),
    "A path-taking MP4 method is supported");
Check(ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputAsync), typeof(string)), "mp4"),
    "A generic path-taking method is supported");
Check(!ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputWavFile), typeof(string)), "mp4"),
    "MP4 must not invoke a WAV export method");
Check(!ExportInvocationPolicy.MatchesFormat("OutputExoCommand", "mp4"),
    "MP4 must not invoke an EXO command");
Check(ExportInvocationPolicy.MatchesFormat("OutputMovie", "mp4"),
    "A generic movie exporter remains available for MP4");
Check(ExportInvocationPolicy.HasDialogOnlyMethod(
    [Method(nameof(ExportMethods.OutputExoFile)), Method(nameof(ExportMethods.OutputMp4File))], "mp4"),
    "A matching parameterless method requires a dialog");
Check(!ExportInvocationPolicy.HasDialogOnlyMethod([Method(nameof(ExportMethods.OutputExoFile))], "mp4"),
    "An EXO-only method is unavailable for MP4");
Console.WriteLine("Export invocation policy tests passed");

class ExportMethods
{
    public void OutputExoFile() { }
    public void OutputMp4File() { }
    public void OutputMp4File(string path) { }
    public void OutputWavFile(string path) { }
    public void OutputAsync(string path) { }
}
