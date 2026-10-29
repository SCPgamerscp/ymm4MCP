using System;
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
}
