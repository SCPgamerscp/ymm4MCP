using System;
using System.IO;
using System.Text.RegularExpressions;

namespace YMM4McpPlugin
{
    internal static class JobSnapshotSafety
    {
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
                File.Move(temp, path, overwrite: true);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); }
                catch { /* Preserve the original write error. */ }
            }
        }
    }
}
