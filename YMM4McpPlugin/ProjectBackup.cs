using System.IO;

namespace YMM4McpPlugin;

internal static class ProjectBackup
{
    internal static string? BeforeOverwrite(string sourcePath, string backupDirectory)
    {
        if (!File.Exists(sourcePath)) return null;
        Directory.CreateDirectory(backupDirectory);
        string backupPath = Path.Combine(backupDirectory,
            "save_" + DateTime.UtcNow.ToString("yyyyMMddTHHmmssfff") + "_" + Guid.NewGuid().ToString("N") + ".ymmp");
        // Never publish a partially copied backup as a restorable .ymmp file.
        // Stage in the same directory so the final rename is atomic.
        string stagingPath = backupPath + ".partial";
        try
        {
            // Hold a read-only sharing lease for the entire copy. File.Copy alone
            // can snapshot a project while another writer is modifying it,
            // yielding a formally complete but internally inconsistent backup.
            using (var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // A zero-byte .ymmp is not a restorable project. Never publish it
                // as a valid recovery point.
                if (source.Length == 0)
                    throw new InvalidDataException("Cannot back up an empty YMM4 project.");
                using var staging = new FileStream(stagingPath, FileMode.CreateNew,
                    FileAccess.Write, FileShare.None);
                source.CopyTo(staging);
                staging.Flush(flushToDisk: true);
            }
            File.Move(stagingPath, backupPath);
            return backupPath;
        }
        catch
        {
            try { File.Delete(stagingPath); }
            catch (IOException) { /* Preserve the original backup failure. */ }
            catch (UnauthorizedAccessException) { /* Preserve the original backup failure. */ }
            throw;
        }
    }
}
