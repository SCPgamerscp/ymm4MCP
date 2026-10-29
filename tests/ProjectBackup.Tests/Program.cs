using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

string root = Path.Combine(Path.GetTempPath(), "ymm4-backup-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    string source = Path.Combine(root, "project.ymmp");
    string backups = Path.Combine(root, "backups");
    Check(ProjectBackup.BeforeOverwrite(source, backups) == null,
        "missing project must not create a backup");
    File.WriteAllText(source, "original project");
    string? snapshot = ProjectBackup.BeforeOverwrite(source, backups);
    Check(snapshot != null && File.Exists(snapshot) && File.ReadAllText(snapshot) == "original project",
        "a completed backup must contain the original project bytes");
    Check(Directory.GetFiles(backups, "*.partial").Length == 0,
        "a successful backup must not leave a partial file");
    File.WriteAllText(source, "modified project");
    Check(File.ReadAllText(snapshot!) == "original project",
        "later project edits must not mutate the recovery snapshot");

    File.WriteAllBytes(source, Array.Empty<byte>());
    bool emptyRejected = false;
    try { ProjectBackup.BeforeOverwrite(source, backups); }
    catch (InvalidDataException) { emptyRejected = true; }
    Check(emptyRejected, "empty projects must not be published as restorable backups");
    Check(Directory.GetFiles(backups, "*.ymmp").Length == 1
        && Directory.GetFiles(backups, "*.partial").Length == 0,
        "a rejected empty backup must not leave published or staged files");

    // An existing exclusive writer must prevent a potentially inconsistent
    // snapshot; never silently copy a project being actively overwritten.
    File.WriteAllText(source, "actively edited");
    using (var writer = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
    {
        bool lockedRejected = false;
        try { ProjectBackup.BeforeOverwrite(source, backups); }
        catch (IOException) { lockedRejected = true; }
        Check(lockedRejected, "an exclusively locked project must not be backed up");
    }
    Check(Directory.GetFiles(backups, "*.ymmp").Length == 1
        && Directory.GetFiles(backups, "*.partial").Length == 0,
        "a locked source must not leave a misleading recovery point");
}
finally { Directory.Delete(root, recursive: true); }

Console.WriteLine("Project backup integrity tests passed");
