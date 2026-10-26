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
            File.Copy(sourcePath, stagingPath);
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
