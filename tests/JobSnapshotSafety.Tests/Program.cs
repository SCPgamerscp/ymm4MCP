using System.Text.Json;
using YMM4McpPlugin;

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
