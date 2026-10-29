using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace YMM4McpPlugin
{
    internal static class JobSnapshotSafety
    {
        // File.Move(overwrite: true) is not safe against concurrent writers on
        // Windows: even distinct temporary files can race at the destination.
        // Serialize only the final replacement, not the temporary file write.
        private static readonly object SnapshotReplaceLock = new();

        private static readonly Regex JobId = new(
            @"^job_[A-Za-z0-9]{8,64}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        internal static bool IsRestorable(string? id, string? kind, string? status, string? key)
        {
            if (id == null || !JobId.IsMatch(id) || kind != "export")
                return false;
            if (status is not ("queued" or "running" or "completed" or "failed" or "cancelled" or "interrupted"))
                return false;
            return key == null || ExportInvocationPolicy.IsValidIdempotencyKey(key);
        }

        internal static void WriteAtomic(string path, string content)
        {
            // Unique temp files prevent a previous failed write from corrupting
            // the next snapshot. The caller serializes snapshot generation/writes.
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temp, content);
                lock (SnapshotReplaceLock)
                {
                    // A static lock only protects threads in this YMM4 process.
                    // Two YMM4 processes can still write the same per-user job
                    // snapshot, so use a stable per-path OS mutex as well.
                    string fullPath = Path.GetFullPath(path);
                    if (OperatingSystem.IsWindows())
                        fullPath = fullPath.ToUpperInvariant();
                    string mutexName = "YMM4McpJobSnapshot_" +
                        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)));
                    using var mutex = new Mutex(false, mutexName);
                    bool acquired = false;
                    try
                    {
                        try { acquired = mutex.WaitOne(TimeSpan.FromSeconds(30)); }
                        catch (AbandonedMutexException) { acquired = true; }
                        if (!acquired)
                            throw new IOException("Timed out waiting to replace the job snapshot.");
                        File.Move(temp, path, overwrite: true);
                    }
                    finally
                    {
                        if (acquired)
                            mutex.ReleaseMutex();
                    }
                }
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); }
                catch { /* Preserve the original write error. */ }
            }
        }
    }
}
