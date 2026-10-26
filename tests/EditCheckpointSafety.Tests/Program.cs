using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var seen = new HashSet<string>(StringComparer.Ordinal);
EditCheckpointSafety.ValidateBinding("scene-1", "native:item-1", seen);
Check(seen.Count == 1, "valid binding rejected");

void ExpectInvalid(string logicalId, string itemId)
{
    try
    {
        EditCheckpointSafety.ValidateBinding(logicalId, itemId, seen);
        throw new Exception("invalid binding accepted");
    }
    catch (ArgumentException) { }
}

ExpectInvalid("scene-1", "native:other"); // duplicate logical ID
ExpectInvalid(" ", "native:other");
ExpectInvalid(" scene-2", "native:other");
ExpectInvalid("scene\n3", "native:other");
ExpectInvalid("scene-4", "native:\0bad");
ExpectInvalid(new string('x', 513), "native:other");
Check(EditCheckpointSafety.MaxBindings == 5000, "binding limit changed unexpectedly");

string root = Path.Combine(Path.GetTempPath(), "ymm4-checkpoint-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    string source = Path.Combine(root, "source.ymmp");
    string destination = Path.Combine(root, "cp.ymmp");
    File.WriteAllBytes(source, new byte[] { 1, 2, 3, 4, 5 });
    EditCheckpointSafety.CopyBackupAtomically(source, destination);
    Check(File.ReadAllBytes(destination).SequenceEqual(File.ReadAllBytes(source)),
        "published checkpoint bytes differ from source");
    Check(!Directory.EnumerateFiles(root, "*.tmp").Any(), "temporary file leaked after success");

    string failed = Path.Combine(root, "missing-backup.ymmp");
    try
    {
        EditCheckpointSafety.CopyBackupAtomically(Path.Combine(root, "not-found.ymmp"), failed);
        throw new Exception("copying a missing project unexpectedly succeeded");
    }
    catch (FileNotFoundException) { }
    Check(!File.Exists(failed), "failed copy published a backup file");
    Check(!Directory.EnumerateFiles(root, "*.tmp").Any(), "temporary file leaked after failure");
}
finally
{
    Directory.Delete(root, recursive: true);
}
Console.WriteLine("Edit checkpoint safety tests passed");
