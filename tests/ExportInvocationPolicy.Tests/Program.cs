using System.Reflection;
using System.Windows.Input;
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

Check(!ExportInvocationPolicy.ReturnsCompletionTask(Method(nameof(ExportMethods.OutputMp4File), typeof(string))),
    "A path-taking void method cannot be observed through completion");
Check(ExportInvocationPolicy.ReturnsCompletionTask(Method(nameof(ExportMethods.OutputAwaitable), typeof(string))),
    "A native ValueTask<T> export must be supported");
Check(ExportInvocationPolicy.GetCompletionTask(new ValueTask<int>(42)) is Task<int> converted &&
      await converted == 42, "The original ValueTask<T> must be converted into an observable Task");
Check(ExportInvocationPolicy.GetCompletionTask(null) is null, "A null return must not imply completion");

var plain = new PlainExportCommand();
Check(ExportInvocationPolicy.FindAwaitableCommand(plain) is null && !plain.Executed,
    "A command offering only void Execute must not be invoked");
var asyncCommand = new AsyncExportCommand();
var execute = ExportInvocationPolicy.FindAwaitableCommand(asyncCommand);
var nativeTask = execute == null ? null
    : ExportInvocationPolicy.InvokeProjectPathMethod(execute, asyncCommand, "episode.ymmp");
Check(nativeTask == asyncCommand.Completion.Task && !asyncCommand.Executed &&
      asyncCommand.OutputPath == "episode.ymmp" && nativeTask is { IsCompleted: false },
    "A command must use ExecuteAsync with the requested path and track its actual pending task");
asyncCommand.Completion.SetResult(true);
await nativeTask!;

var project = new ProjectFileMethods();
string projectPath = "example.ymmp";
var openMethod = typeof(ProjectFileMethods).GetMethod(nameof(ProjectFileMethods.OpenProject))!;
var opening = ExportInvocationPolicy.InvokeProjectPathMethod(openMethod, project, projectPath);
Check(opening == project.Opening.Task && opening is { IsCompleted: false } && project.LastPath == projectPath,
    "Opening a project must preserve the pending ValueTask completion");
project.Opening.SetResult(true);
await opening!;
var saveMethod = typeof(ProjectFileMethods).GetMethod(nameof(ProjectFileMethods.SaveProjectAs))!;
Check(ExportInvocationPolicy.InvokeProjectPathMethod(saveMethod, project, "copy.ymmp") is null &&
      project.LastPath == "copy.ymmp", "Synchronous project methods must remain usable");
var brokenMethod = typeof(ProjectFileMethods).GetMethod(nameof(ProjectFileMethods.OpenWithoutCompletion))!;
bool rejectedMissingTask = false;
try { ExportInvocationPolicy.InvokeProjectPathMethod(brokenMethod, project, projectPath); }
catch (InvalidOperationException) { rejectedMissingTask = true; }
Check(rejectedMissingTask, "An asynchronous project method without a task must not be acknowledged as complete");

// Regression: export jobs persist idempotency keys, so embedded control characters
// must be rejected before the request is recorded or replayed.
Check(ExportInvocationPolicy.IsValidIdempotencyKey(""), "Omitted key remains valid");
Check(ExportInvocationPolicy.IsValidIdempotencyKey(new string('a', 128)), "128-character key must be accepted");
Check(ExportInvocationPolicy.IsValidIdempotencyKey("編集 job-1"), "Normal Unicode and internal spaces remain valid");
foreach (var key in new[] { " a", "a ", new string('a', 129), "a\nb", "a\rb", "a\tb", "a\0b", "a\u007fb", "a\u001fb" })
    Check(!ExportInvocationPolicy.IsValidIdempotencyKey(key), "Invalid idempotency key accepted: " + Convert.ToHexString(System.Text.Encoding.UTF8.GetBytes(key)));

string outputDirectory = Directory.CreateTempSubdirectory("ymm4-export-destination-").FullName;
try
{
    string output = Path.Combine(outputDirectory, "episode.mp4");
    Check(ExportInvocationPolicy.TryCaptureOutputBaseline(output, false, out var missing) && missing is null,
        "A new destination may be used without overwrite");
    File.WriteAllBytes(output, [1, 2, 3]);
    Check(!ExportInvocationPolicy.TryCaptureOutputBaseline(output, false, out var blocked) && blocked is null,
        "A destination created while the job waited or was interrupted must block export");
    Check(File.ReadAllBytes(output).SequenceEqual(new byte[] { 1, 2, 3 }),
        "The no-overwrite check must leave the existing file as-is");
    Check(ExportInvocationPolicy.TryCaptureOutputBaseline(output, true, out var baseline) &&
        baseline == (3L, File.GetLastWriteTimeUtc(output)),
        "An allowed overwrite must retain the previous file's length and timestamp for verification");
    string folder = Directory.CreateDirectory(Path.Combine(outputDirectory, "folder.mp4")).FullName;
    Check(!ExportInvocationPolicy.TryCaptureOutputBaseline(folder, true, out _),
        "An existing directory cannot serve as an export file even with overwrite enabled");
}
finally { Directory.Delete(outputDirectory, recursive: true); }

var stability = new ExportOutputStability();
var outputTime = new DateTime(2026, 11, 28, 0, 0, 0, DateTimeKind.Utc);
for (int sample = 0; sample < 5; sample++)
    Check(!stability.Observe((128, outputTime.AddSeconds(sample))),
        "Same-size rewrites must keep the export unstable");
var settled = (128L, outputTime.AddSeconds(4));
Check(!stability.Observe(settled) && !stability.Observe(settled) && stability.Observe(settled),
    "An unchanged export can be verified after three matching polls");
Check(!stability.Observe(null) && !stability.Observe(settled),
    "A disappearance or discarded verification must restart the stability window");

Console.WriteLine("Export invocation policy tests passed");

class ExportMethods
{
    public void OutputExoFile() { }
    public void OutputMp4File() { }
    public void OutputMp4File(string path) { }
    public void OutputWavFile(string path) { }
    public void OutputAsync(string path) { }
    public void OutputWithFormat(string path, string format) { }
    public ValueTask<int> OutputAwaitable(string path) => new(42);
    public void OutputWithOutputFormat(string path, string outputFormat) { }
    public void OutputWithPreset(string path, string presetPath) { }
    public void OutputWithDirectory(string path, string outputDirectory) { }
}

class ProjectFileMethods
{
    public TaskCompletionSource<bool> Opening { get; } = new();
    public string? LastPath { get; private set; }

    public ValueTask OpenProject(string path)
    {
        LastPath = path;
        return new ValueTask(Opening.Task);
    }

    public void SaveProjectAs(string path) => LastPath = path;
    public Task? OpenWithoutCompletion(string path) => null;
}

class PlainExportCommand : ICommand
{
    public bool Executed { get; private set; }
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => Executed = true;
    public event EventHandler? CanExecuteChanged { add { } remove { } }
}

class AsyncExportCommand : PlainExportCommand
{
    public TaskCompletionSource<bool> Completion { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public string? OutputPath { get; private set; }

    public Task ExecuteAsync(object? parameter)
    {
        OutputPath = parameter as string;
        return Completion.Task;
    }
}
