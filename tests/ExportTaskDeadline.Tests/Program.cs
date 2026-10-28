using YMM4McpPlugin;

static void Check(bool value, string message)
{
    if (!value) throw new Exception(message);
}

Check(await ExportTaskDeadline.WaitForCompletion(
    Task.CompletedTask, DateTime.UtcNow.AddSeconds(2), CancellationToken.None),
    "A completed native export must succeed");

var never = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
Check(!await ExportTaskDeadline.WaitForCompletion(
    never.Task, DateTime.UtcNow.AddMilliseconds(-1), CancellationToken.None),
    "An already-expired deadline must not await the native task");
Check(!await ExportTaskDeadline.WaitForCompletion(
    never.Task, DateTime.UtcNow.AddMilliseconds(50), CancellationToken.None),
    "A hung native export must time out");

bool failed = false;
try
{
    await ExportTaskDeadline.WaitForCompletion(
        Task.FromException(new InvalidOperationException("native encoder failed")),
        DateTime.UtcNow.AddSeconds(2), CancellationToken.None);
}
catch (InvalidOperationException) { failed = true; }
Check(failed, "Native exporter failures must be propagated");

using var cancellation = new CancellationTokenSource();
cancellation.CancelAfter(20);
bool cancelled = false;
try
{
    await ExportTaskDeadline.WaitForCompletion(
        never.Task, DateTime.UtcNow.AddSeconds(2), cancellation.Token);
}
catch (OperationCanceledException) { cancelled = true; }
Check(cancelled, "Cancellation must interrupt the native export wait");

Console.WriteLine("Export task deadline regression tests passed");
