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

// A corrupt or legacy edits.json must not bypass validation during restart.
var stored = new (string? Id, string? ItemId)[]
{
    ("scene-a", "native:a"), ("scene-b", "native:b")
};
var storedHash = new string('a', 64);
Check(EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash, stored),
    "valid persisted EditPlan bindings rejected");
Check(EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash,
        Array.Empty<(string? Id, string? ItemId)>()),
    "empty persisted bindings must remain valid");
Check(!EditCheckpointSafety.ValidStoredBindingRecord(null, storedHash, stored),
    "missing idempotency key accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", "short", stored),
    "malformed persisted plan hash accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash, null),
    "missing persisted item list accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash,
        new (string? Id, string? ItemId)[] { ("same", "native:a"), ("same", "native:b") }),
    "duplicate logical IDs in persisted bindings accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash,
        new (string? Id, string? ItemId)[] { ("a", "native:a"), ("b", "native:a") }),
    "duplicate native IDs in persisted bindings accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash,
        new (string? Id, string? ItemId)[] { ("a", null) }),
    "null persisted native ID accepted");
Check(!EditCheckpointSafety.ValidStoredBindingRecord("plan-1", storedHash,
        Enumerable.Range(0, EditCheckpointSafety.MaxBindings + 1)
            .Select(i => ((string?)$"scene-{i}", (string?)$"native:{i}"))),
    "oversized persisted binding list accepted");

// A first binding save is subject to the same live timeline checks as a retry.
// No previous record must not mean arbitrary or nonexistent item IDs are accepted.
var initial = Array.Empty<(string Id, string ItemId, string? Revision)>();
var initialValid = new[] { (Id: "scene-a", ItemId: "native:a", Revision: (string?)"rev-a") };
Check(EditCheckpointSafety.CompatibleBindingUpdate(initial, initialValid,
        new[] { "native:a" }), "first binding save for a live item rejected");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(initial, initialValid,
        new[] { "native:other" }), "first save accepted a nonexistent native item");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(initial, initialValid,
        new[] { "native:a", "" }), "first save accepted a timeline with an unidentified item");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(initial, initialValid,
        new[] { "native:a", "native:a" }), "first save accepted duplicate live identities");

// Reconcile must extend partial plans without allowing live native bindings to
// be silently remapped or revised. A deleted native item may be replaced.
var original = new (string Id, string ItemId, string? Revision)[]
{
    ("scene-a", "native:a", "rev-a")
};
var expanded = new (string Id, string ItemId, string? Revision)[]
{
    ("scene-a", "native:a", "rev-a"),
    ("scene-b", "native:b", "rev-b")
};
Check(EditCheckpointSafety.CompatibleBindingUpdate(original, expanded,
        new[] { "native:a", "native:b" }),
    "partial EditPlan binding extension must be allowed");
Check(EditCheckpointSafety.CompatibleBindingUpdate(expanded,
        new[] { expanded[1], expanded[0] }, new[] { "native:a", "native:b" }),
    "reordering identical bindings must not cause a conflict");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original,
        new[] { ("scene-a", "native:b", "rev-b") }, new[] { "native:a", "native:b" }),
    "a surviving native item must not be remapped");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original,
        new[] { ("scene-a", "native:a", "changed") }, new[] { "native:a" }),
    "a surviving native item revision must not be changed");
Check(EditCheckpointSafety.CompatibleBindingUpdate(original,
        new[] { ("scene-a", "native:b", "rev-b") }, new[] { "native:b" }),
    "a deleted native item may be replaced during reconcile");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original,
        new[] { ("scene-a", "native:missing", "rev-b") }, new[] { "native:b" }),
    "replacement bindings must point to live native items");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original, expanded,
        new[] { "native:a", "native:a", "native:b" }),
    "ambiguous live item IDs must fail closed");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original, expanded, null),
    "missing timeline snapshot must not authorize remaps");
Check(!EditCheckpointSafety.CompatibleBindingUpdate(original,
        new[] { ("scene-a", "native:a", "rev-a"), ("scene-b", "native:a", "rev-b") },
        new[] { "native:a" }),
    "two logical bindings must not claim the same native item");

// A missing or ambiguous snapshot must never authorize destructive rollback.
EditCheckpointSafety.ValidateSnapshotItemIds(Array.Empty<string>());
EditCheckpointSafety.ValidateSnapshotItemIds(new[] { "item-1", "item-2" });
foreach (var invalid in new[]
{
    new[] { "item-1", "item-1" },
    new[] { "item-1", "" },
    new[] { "item-1", " item-2" },
    new[] { "item-1", "item\n2" },
})
{
    try
    {
        EditCheckpointSafety.ValidateSnapshotItemIds(invalid);
        throw new Exception("ambiguous checkpoint snapshot accepted");
    }
    catch (ArgumentException) { }
}
Check(!EditCheckpointSafety.SameProject(null, "C:\\project.ymmp"),
    "unknown checkpoint project must fail closed");
Check(!EditCheckpointSafety.SameProject("C:\\one.ymmp", null),
    "unknown current project must fail closed");
Check(!EditCheckpointSafety.SameProject("C:\\one.ymmp", "C:\\two.ymmp"),
    "different projects must not share checkpoints");
Check(EditCheckpointSafety.SameProject("C:\\one.ymmp", "C:\\ONE.ymmp"),
    "Windows project path comparison should be case-insensitive");

// A checkpoint backup is only valid while the same project remains saved.
// A project switch or an edit during the copy must invalidate the backup.
Check(EditCheckpointSafety.IsStableSavedProject("C:\\video\\one.ymmp",
        "C:\\VIDEO\\ONE.ymmp", true), "stable saved project rejected");
Check(!EditCheckpointSafety.IsStableSavedProject("C:\\video\\one.ymmp",
        "C:\\video\\two.ymmp", true), "project switch accepted during backup");
Check(!EditCheckpointSafety.IsStableSavedProject("C:\\video\\one.ymmp",
        "C:\\video\\one.ymmp", false), "unsaved edit accepted during backup");
Check(!EditCheckpointSafety.IsStableSavedProject("C:\\video\\one.ymmp",
        "C:\\video\\one.ymmp", null), "unknown save state accepted during backup");
Check(!EditCheckpointSafety.IsStableSavedProject("C:\\video\\one.ymmp",
        null, true), "unknown project accepted during backup");

string ownedDir = Path.Combine(Path.GetTempPath(), "ymm4-owned-backups");
Check(EditCheckpointSafety.IsOwnedBackupPath(Path.Combine(ownedDir, "cp_12345678.ymmp"), ownedDir),
    "owned checkpoint backup rejected");
Check(!EditCheckpointSafety.IsOwnedBackupPath(Path.Combine(ownedDir, "..", "outside", "cp_12345678.ymmp"), ownedDir),
    "backup cleanup escaped its directory");
Check(!EditCheckpointSafety.IsOwnedBackupPath(Path.Combine(ownedDir, "other.ymmp"), ownedDir),
    "non-checkpoint backup accepted for deletion");
Check(!EditCheckpointSafety.IsOwnedBackupPath(Path.Combine(ownedDir, "cp_12345678.txt"), ownedDir),
    "non-ymmp file accepted for deletion");
Check(!EditCheckpointSafety.IsOwnedBackupPath(null, ownedDir),
    "null backup path accepted for deletion");

// Never delete a checkpoint backup that is currently open as a project.
string pruningDir = Path.Combine(Path.GetTempPath(), "ymm4-prune-" + Guid.NewGuid().ToString("N"));
string activeBackup = Path.Combine(pruningDir, "cp_12345678.ymmp");
Check(!EditCheckpointSafety.CanPruneBackup(activeBackup, pruningDir, activeBackup),
    "active restored project must not be pruned");
Check(!EditCheckpointSafety.CanPruneBackup(activeBackup, pruningDir, null),
    "unknown current project must not authorize backup deletion");
Check(EditCheckpointSafety.CanPruneBackup(activeBackup, pruningDir, "C:\\other\\project.ymmp"),
    "inactive owned backup should be prunable");
Check(!EditCheckpointSafety.CanPruneBackup(Path.Combine(pruningDir, "..", "outside", "cp_12345678.ymmp"),
        pruningDir, "C:\\other\\project.ymmp"),
    "backup pruning must reject escaped paths");

// Full rollback requires an owned backup with the exact SHA-256 recorded at
// checkpoint creation. Replacing the bytes while preserving size/timestamps
// must never result in a successful restore.
string restoreDir = Path.Combine(Path.GetTempPath(), "ymm4-restore-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(restoreDir);
try
{
    const string checkpointId = "cp_12345678";
    const string project = "C:\\video\\project.ymmp";
    string backup = Path.Combine(restoreDir, checkpointId + ".ymmp");
    File.WriteAllText(backup, "original backup payload");
    string hash = EditCheckpointSafety.ComputeBackupSha256(backup);
    Check(EditCheckpointSafety.IsValidBackupSha256(hash), "generated SHA-256 is invalid");
    Check(!EditCheckpointSafety.IsValidBackupSha256(null), "missing digest accepted");
    Check(!EditCheckpointSafety.IsValidBackupSha256(new string('a', 63)), "short digest accepted");
    Check(!EditCheckpointSafety.IsValidBackupSha256(new string('g', 64)), "nonhex digest accepted");
    Check(EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project, project, hash),
        "valid owned hashed backup must be restorable");
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project, project, null),
        "legacy backup without integrity digest must fail closed");
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, "cp_87654321", project, project, hash),
        "a checkpoint must not restore another checkpoint's backup");
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project,
            "C:\\video\\different.ymmp", hash), "backup must not restore into a different project");
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, null, project, hash),
        "unknown source project must reject restore");
    Check(!EditCheckpointSafety.IsRestorableBackup(
            Path.Combine(restoreDir, "..", "outside", checkpointId + ".ymmp"),
            restoreDir, checkpointId, project, project, hash),
        "restore must reject paths outside the private backup directory");
    var originalTimestamp = File.GetLastWriteTimeUtc(backup);
    File.WriteAllText(backup, "modified backup payload");
    File.SetLastWriteTimeUtc(backup, originalTimestamp);
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project, project, hash),
        "modified backup with identical length and timestamp must fail");
    File.WriteAllText(backup, "");
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project, project, hash),
        "empty backups must not be opened");
    File.Delete(backup);
    Check(!EditCheckpointSafety.IsRestorableBackup(backup, restoreDir, checkpointId, project, project, hash),
        "missing backups must not be opened");
}
finally
{
    Directory.Delete(restoreDir, recursive: true);
}

string root = Path.Combine(Path.GetTempPath(), "ymm4-checkpoint-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    string source = Path.Combine(root, "source.ymmp");
    string destination = Path.Combine(root, "cp.ymmp");
    File.WriteAllBytes(source, new byte[] { 1, 2, 3, 4, 5 });
    string copiedHash = EditCheckpointSafety.CopyBackupVerified(source, destination);
    Check(File.ReadAllBytes(destination).SequenceEqual(File.ReadAllBytes(source)),
        "published checkpoint bytes differ from source");
    Check(string.Equals(copiedHash, EditCheckpointSafety.ComputeBackupSha256(source), StringComparison.Ordinal),
        "verified copy did not return the actual source digest");
    Check(!Directory.EnumerateFiles(root, "*.tmp").Any(), "temporary file leaked after success");

    string failed = Path.Combine(root, "missing-backup.ymmp");
    try
    {
        EditCheckpointSafety.CopyBackupVerified(Path.Combine(root, "not-found.ymmp"), failed);
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
