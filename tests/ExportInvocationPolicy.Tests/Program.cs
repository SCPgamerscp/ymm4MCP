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
Check(ExportInvocationPolicy.MatchesFormat("OutputSavingMp4", "mp4"),
    "Saving must not be misclassified as AVI");
Check(ExportInvocationPolicy.MatchesFormat("OutputSavingWav", "wav"),
    "Saving must not be misclassified as AVI for WAV");
Check(ExportInvocationPolicy.MatchesFormat("OutputWavingMp4", "mp4"),
    "Waving must not be misclassified as WAV");
Check(!ExportInvocationPolicy.MatchesFormat("OutputAVIFile", "mp4"),
    "An explicit AVI exporter must not be used for MP4");
Check(!ExportInvocationPolicy.MatchesFormat("OutputWavFile", "mp4"),
    "An explicit WAV exporter must not be used for MP4");
Check(!ExportInvocationPolicy.MatchesFormat("OutputMovie", "wav"),
    "A generic video exporter must not be used for WAV");
Check(ExportInvocationPolicy.MatchesFormat("OutputMovieWav", "wav"),
    "An explicitly WAV-capable movie exporter remains available");
Check(ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputWithFormat), typeof(string), typeof(string)), "mp4"),
    "A path plus format method must remain supported");
Check(ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputWithOutputFormat), typeof(string), typeof(string)), "mp4"),
    "A path plus outputFormat method must remain supported");
Check(!ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputWithPreset), typeof(string), typeof(string)), "mp4"),
    "A second preset path must not receive the format string");
Check(!ExportInvocationPolicy.AcceptsPath(Method(nameof(ExportMethods.OutputWithDirectory), typeof(string), typeof(string)), "mp4"),
    "A second output directory must not receive the format string");

Check(ExportInvocationPolicy.HasDialogOnlyMethod(
    [Method(nameof(ExportMethods.OutputExoFile)), Method(nameof(ExportMethods.OutputMp4File))], "mp4"),
    "A matching parameterless method requires a dialog");
Check(!ExportInvocationPolicy.HasDialogOnlyMethod([Method(nameof(ExportMethods.OutputExoFile))], "mp4"),
    "An EXO-only method is unavailable for MP4");
// Regression: export jobs persist idempotency keys, so embedded control characters
// must be rejected before the request is recorded or replayed.
Check(ExportInvocationPolicy.IsValidIdempotencyKey(""), "Omitted key remains valid");
Check(ExportInvocationPolicy.IsValidIdempotencyKey(new string('a', 128)), "128-character key must be accepted");
Check(ExportInvocationPolicy.IsValidIdempotencyKey("編集 job-1"), "Normal Unicode and internal spaces remain valid");
foreach (var key in new[] { " a", "a ", new string('a', 129), "a\nb", "a\rb", "a\tb", "a\0b", "a\u007fb", "a\u001fb" })
    Check(!ExportInvocationPolicy.IsValidIdempotencyKey(key), "Invalid idempotency key accepted: " + Convert.ToHexString(System.Text.Encoding.UTF8.GetBytes(key)));
Console.WriteLine("Export invocation policy tests passed");

class ExportMethods
{
    public void OutputExoFile() { }
    public void OutputMp4File() { }
    public void OutputMp4File(string path) { }
    public void OutputWavFile(string path) { }
    public void OutputAsync(string path) { }
    public void OutputWithFormat(string path, string format) { }
    public void OutputWithOutputFormat(string path, string outputFormat) { }
    public void OutputWithPreset(string path, string presetPath) { }
    public void OutputWithDirectory(string path, string outputDirectory) { }
}
