using System.Net;

namespace YMM4McpPlugin
{
    // Compile the production job registry without loading the YMM4 UI.
    public partial class McpHttpServer
    {
        private readonly SemaphoreSlim _editGate = new(1);

        private static object Failure(string code, string message) => new { code, message };
        private void Log(string message) { }
        private Task<object> EnqueueExport(HttpListenerRequest request) => throw new NotSupportedException();
        private object RestartExport(McpJob job) => throw new NotSupportedException();
        private static bool SameJobRequest(McpJob job, string kind, Dictionary<string, object?> request,
            ExportProjectIdentity source) => job.Kind == kind && job.ExportSource == source
                && job.Request["output_path"]?.ToString() == request["output_path"]?.ToString();

        internal static void CheckAcceptance(string directory, ExportProjectIdentity source, Action<bool, string> check)
        {
            McpSettings.DirectoryPath = directory;
            var registry = new McpHttpServer();
            var request = new Dictionary<string, object?> { ["output_path"] = Path.Combine(directory, "export.mp4") };

            // A directory with the destination name rejects the atomic file replacement.
            Directory.CreateDirectory(JobsFilePath);
            var rejected = registry.CreateJob("export", request, "export-test", source,
                out bool created, out bool conflict);
            check(rejected == null && !created && !conflict && registry.FindIdempotentJob("export-test") == null,
                "a job whose first snapshot failed must not be registered or acknowledged");
            Directory.Delete(JobsFilePath);

            // Start no worker. The accepted job must already survive an immediate restart.
            var accepted = registry.CreateJob("export", request, "export-test", source, out created, out conflict);
            check(accepted != null && created && !conflict, "the same key must remain retryable after a failed write");
            var restored = new McpHttpServer();
            restored.RestorePersistedJobs();
            var pending = restored.FindIdempotentJob("export-test");
            check(pending != null && pending.Id == accepted!.Id && pending.Status == "interrupted",
                "the accepted job ID must be recoverable before the export worker runs");

            var retried = registry.CreateJob("export", request, "export-test", source, out created, out conflict);
            check(ReferenceEquals(retried, accepted) && !created && !conflict,
                "retrying an accepted request must return the original job");
        }
    }

    internal static class McpSettings
    {
        public static string DirectoryPath { get; set; } = "";
        public static void PrepareDirectory() => Directory.CreateDirectory(DirectoryPath);
    }
}
