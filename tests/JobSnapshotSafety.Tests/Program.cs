using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using YMM4McpPlugin;

// Child mode exercises the production snapshot writer from a different OS
// process; a C# static lock cannot coordinate these processes.
if (args.Length == 3 && args[0] == "--snapshot-writer")
{
    int count = int.Parse(args[2]);
    for (int i = 0; i < count; i++)
        JobSnapshotSafety.WriteAtomic(args[1],
            "{\"writer\":" + Environment.ProcessId + ",\"sequence\":" + i + "}");
    return;
}

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

foreach (var status in new[] { "queued", "running", "completed", "failed", "cancelled", "interrupted" })
    Check(JobSnapshotSafety.IsRestorable("job_12345678", "export", status, "valid-key"),
        "Valid persisted export job was rejected: " + status);
Check(!JobSnapshotSafety.IsRestorable(null, "export", "completed", null),
    "Null job IDs must be rejected");
Check(!JobSnapshotSafety.IsRestorable("../jobs.json", "export", "completed", null),
    "Malformed IDs must be rejected");
Check(!JobSnapshotSafety.IsRestorable("job_12345678", "unknown", "completed", null),
    "Unknown job kinds must be rejected");
Check(!JobSnapshotSafety.IsRestorable("job_12345678", "export", "unknown", null),
    "Unknown statuses must be rejected");
Check(!JobSnapshotSafety.IsRestorable("job_12345678", "export", "completed", "bad\nkey"),
    "Malformed idempotency keys must be rejected");

var directory = Path.Combine(Path.GetTempPath(), "ymm4-job-snapshot-test-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(directory);
try
{
    var path = Path.Combine(directory, "jobs.json");
    JobSnapshotSafety.WriteAtomic(path, "{\"version\":1}");
    JobSnapshotSafety.WriteAtomic(path, "{\"version\":2}");
    Check(JsonDocument.Parse(File.ReadAllText(path)).RootElement.GetProperty("version").GetInt32() == 2,
        "New snapshots must atomically replace old snapshots");
    // Repeated parallel replacements used to throw UnauthorizedAccessException
    // on Windows even though each writer had a unique staging filename.
    for (int round = 0; round < 4; round++)
    {
        Parallel.For(0, 32, i => JobSnapshotSafety.WriteAtomic(
            path, "{\"version\":" + (round * 32 + i) + "}"));
    }
    var version = JsonDocument.Parse(File.ReadAllText(path)).RootElement.GetProperty("version").GetInt32();
    Check(version >= 96 && version < 128, "Concurrent snapshot writes must never create mixed JSON");
    Check(!Directory.EnumerateFiles(directory, "*.tmp").Any(),
        "Successful snapshots must not leave temporary files");

    // Multiple YMM4 processes can share the same per-user snapshot directory.
    // The static lock tested above does not protect this scenario.
    var children = new List<Process>();
    try
    {
        for (int worker = 0; worker < 8; worker++)
        {
            var start = new ProcessStartInfo("dotnet")
            {
                UseShellExecute = false,
                RedirectStandardError = true
            };
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
            start.ArgumentList.Add("--snapshot-writer");
            start.ArgumentList.Add(path);
            start.ArgumentList.Add("24");
            children.Add(Process.Start(start) ??
                throw new Exception("Could not launch cross-process snapshot writer"));
        }
        foreach (var child in children)
        {
            if (!child.WaitForExit(90000))
                throw new Exception("Cross-process snapshot writer timed out");
            Check(child.ExitCode == 0,
                "Cross-process snapshot writer failed: " + child.StandardError.ReadToEnd());
        }
        using var snapshot = JsonDocument.Parse(File.ReadAllText(path));
        Check(snapshot.RootElement.TryGetProperty("writer", out _) &&
              snapshot.RootElement.TryGetProperty("sequence", out _),
            "Cross-process writes must leave a complete JSON snapshot");
        Check(!Directory.EnumerateFiles(directory, "*.tmp").Any(),
            "Cross-process snapshot writers must not leak staging files");
    }
    finally
    {
        foreach (var child in children)
        {
            if (!child.HasExited)
                child.Kill(entireProcessTree: true);
            child.Dispose();
        }
    }

    var destinationDirectory = Path.Combine(directory, "occupied");
    Directory.CreateDirectory(destinationDirectory);
    bool failed = false;
    try { JobSnapshotSafety.WriteAtomic(destinationDirectory, "test"); }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
        // Windows may report a directory destination as access denied.
        failed = true;
    }
    Check(failed, "Replacing a directory must fail");
    Check(!Directory.EnumerateFiles(directory, "*.tmp").Any(),
        "Failed snapshots must clean up temporary files");
}
finally
{
    Directory.Delete(directory, recursive: true);
}
Console.WriteLine("Job snapshot safety regression tests passed");
