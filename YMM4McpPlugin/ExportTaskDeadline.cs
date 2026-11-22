using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace YMM4McpPlugin
{
    internal static class ExportTaskDeadline
    {
        // The native export Task may never complete. Cancellation and timeout
        // must both interrupt the MCP job's wait without blocking the edit lock.
        internal static async Task<bool> WaitForCompletion(
            Task nativeTask, DateTime deadlineUtc, CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(nativeTask);
            cancellationToken.ThrowIfCancellationRequested();
            var remaining = deadlineUtc - DateTime.UtcNow;
            if (remaining <= TimeSpan.Zero) return false;

            using var timerCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var timer = Task.Delay(remaining, timerCancellation.Token);
            var completed = await Task.WhenAny(nativeTask, timer);
            cancellationToken.ThrowIfCancellationRequested();
            if (completed != nativeTask) return false;

            timerCancellation.Cancel();
            await nativeTask; // Propagate native exporter errors to the job handler.
            return true;
        }
    }

    // Accessed on the YMM4 dispatcher. Cancelling the MCP wait does not
    // cancel the encoder Task, which may still be using the native exporter.
    internal sealed class NativeExportTaskTracker
    {
        private Task? pending;
        private bool starting;
        internal bool CompletionUnknown { get; private set; }

        internal (bool invoked, string method, string target, Task? task)? TryStart(
            Func<(bool invoked, string method, string target, Task? task)> start)
        {
            if (starting || CompletionUnknown || pending is { IsCompleted: false })
                return null;

            pending = null;
            // A native call can pump the dispatcher before returning a Task.
            starting = true;
            try
            {
                var invocation = start();
                pending = invocation.task;
                CompletionUnknown = invocation.invoked && pending == null;

                // A failure may arrive after the MCP job stopped waiting.
                if (pending is { } task)
                    _ = task.ContinueWith(static completed => { _ = completed.Exception; },
                        CancellationToken.None,
                        TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                        TaskScheduler.Default);

                return invocation;
            }
            catch (TargetInvocationException)
            {
                // The host was entered and may have launched work before throwing.
                CompletionUnknown = true;
                throw;
            }
            finally { starting = false; }
        }
    }
}
