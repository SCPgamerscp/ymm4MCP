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

var tracker = new NativeExportTaskTracker();
var active = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
int invocations = 0;

(bool invoked, string method, string target, Task? task) Start(Task? task)
{
    invocations++;
    return (true, "OutputAsync", "Project", task);
}

var first = tracker.TryStart(() =>
{
    Check(tracker.TryStart(() => Start(null)) is null,
        "A reentrant dispatcher call must not start another exporter during invocation");
    return Start(active.Task);
});
Check(first?.task == active.Task,
    "The first native export must start and expose its Task");
Check(!await ExportTaskDeadline.WaitForCompletion(
    active.Task, DateTime.UtcNow.AddMilliseconds(-1), CancellationToken.None),
    "The MCP job may time out even while the native export is running");
Check(tracker.TryStart(() => Start(null)) is null && invocations == 1,
    "A timed-out wait must not allow another native export while the encoder runs");

using var stopped = new CancellationTokenSource();
stopped.Cancel();
cancelled = false;
try
{
    await ExportTaskDeadline.WaitForCompletion(
        active.Task, DateTime.UtcNow.AddSeconds(2), stopped.Token);
}
catch (OperationCanceledException) { cancelled = true; }
Check(cancelled && tracker.TryStart(() => Start(null)) is null && invocations == 1,
    "A cancelled wait must also leave the native exporter protected");

active.SetResult(true);
var fault = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
Check(tracker.TryStart(() => Start(fault.Task))?.task == fault.Task && invocations == 2,
    "A completed native task must release the exporter for another job");
fault.SetException(new InvalidOperationException("encoder failed after the wait ended"));
Check(tracker.TryStart(() => (false, "", "", null))?.invoked == false,
    "Discovery without invoking an exporter must not reserve it");
Check(tracker.TryStart(() => Start(Task.CompletedTask))?.invoked == true && invocations == 3,
    "A faulted native task must release the exporter, including after unsuccessful discovery");
Check(tracker.TryStart(() => Start(null))?.invoked == true && invocations == 4 &&
      tracker.CompletionUnknown, "An invoked exporter that does not return a Task has an unknown outcome");
Check(tracker.TryStart(() => Start(Task.CompletedTask)) is null && invocations == 4,
    "An unknown native outcome must not allow a second export");

var throwingTracker = new NativeExportTaskTracker();
try
{
    throwingTracker.TryStart(() => throw new System.Reflection.TargetInvocationException(
        new InvalidOperationException("started host then threw")));
    throw new Exception("A synchronous host failure should reach the job handler");
}
catch (System.Reflection.TargetInvocationException) { }
Check(throwingTracker.CompletionUnknown &&
      throwingTracker.TryStart(() => Start(Task.CompletedTask)) is null,
    "A host call that throws before returning its Task must not allow a second export");

Console.WriteLine("Export task deadline regression tests passed");
