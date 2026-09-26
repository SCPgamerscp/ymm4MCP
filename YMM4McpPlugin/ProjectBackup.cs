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
        File.Copy(sourcePath, backupPath);
        return backupPath;
    }
}
