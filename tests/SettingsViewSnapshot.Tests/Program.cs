using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var original = new SettingsViewSnapshot(8765, true, false);
Check(original.Matches(8765, true, false), "unchanged settings should be writable");
Check(!original.Matches(9000, true, false), "stale port update was not detected");
Check(!original.Matches(8765, false, false), "stale auto-start update was not detected");
Check(!original.Matches(8765, true, true), "stale advanced-API update was not detected");

var latest = new SettingsViewSnapshot(9000, false, true);
Check(latest.Matches(9000, false, true), "refresh should allow saving latest values");
Check(!original.Matches(9000, false, true), "old view should still be stale");
// Two concurrent tool windows must not both commit the same stale snapshot.
var current = original;
int writes = 0;
var tasks = Enumerable.Range(0, 24).Select(index => Task.Run(() =>
    SettingsViewSaveCoordinator.TryCommit(original, () => current, () =>
    {
        Thread.Sleep(2); // make a missing compare-and-save gate observable
        current = new SettingsViewSnapshot(9000 + index, true, false);
        Interlocked.Increment(ref writes);
    }))).ToArray();
Task.WaitAll(tasks);
Check(writes == 1, "concurrent stale views both persisted their changes");
Check(tasks.Count(task => task.Result) == 1, "only one stale view may commit");
Check(!SettingsViewSaveCoordinator.TryCommit(original, () => current, () =>
    throw new Exception("stale save must never persist")),
    "stale view unexpectedly passed compare-and-save");
Check(SettingsViewSaveCoordinator.TryCommit(current, () => current, () =>
    Interlocked.Increment(ref writes)), "fresh view should be allowed to save");
Check(writes == 2, "fresh view did not persist");
try
{
    SettingsViewSaveCoordinator.TryCommit(current, () => current,
        () => throw new InvalidOperationException("save failed"));
    throw new Exception("persistence errors must be propagated");
}
catch (InvalidOperationException) { }
// An external YMM4 process can change the on-disk settings without changing
// this process's Server.Settings. A stale file fingerprint must reject a save.
var tempSettings = Path.Combine(Path.GetTempPath(),
    "ymm4-settings-concurrency-" + Guid.NewGuid().ToString("N") + ".json");
try
{
    Check(SettingsViewSaveCoordinator.FileFingerprint(tempSettings) == "<missing>",
        "missing settings file must have a stable sentinel");
    File.WriteAllText(tempSettings, "{\"port\":8765}");
    var observedFile = SettingsViewSaveCoordinator.FileFingerprint(tempSettings);
    Check(!SettingsViewSaveCoordinator.SameReadableFingerprint(
            SettingsViewSaveCoordinator.UnreadableFingerprint,
            SettingsViewSaveCoordinator.UnreadableFingerprint),
        "two unreadable fingerprints must never be considered unchanged");
    Check(!SettingsViewSaveCoordinator.SameReadableFingerprint(
            observedFile, SettingsViewSaveCoordinator.UnreadableFingerprint),
        "a newly unreadable settings file must block saves");
    Check(!SettingsViewSaveCoordinator.SameReadableFingerprint(
            SettingsViewSaveCoordinator.UnreadableFingerprint, observedFile),
        "an initially unreadable settings file must block saves");
    Check(SettingsViewSaveCoordinator.SameReadableFingerprint(observedFile, observedFile),
        "readable unchanged settings should remain writable");
    using (var lockedFile = new FileStream(tempSettings, FileMode.Open,
               FileAccess.ReadWrite, FileShare.None))
    {
        Check(SettingsViewSaveCoordinator.FileFingerprint(tempSettings) ==
              SettingsViewSaveCoordinator.UnreadableFingerprint,
            "locked settings must be classified as unreadable without throwing");
        bool wroteLocked = false;
        Check(!SettingsViewSaveCoordinator.TryCommit(current, () => current,
                () => wroteLocked = true,
                () => SettingsViewSaveCoordinator.SameReadableFingerprint(
                    observedFile, SettingsViewSaveCoordinator.FileFingerprint(tempSettings))),
            "unreadable on-disk settings must block a save");
        Check(!wroteLocked, "locked settings unexpectedly persisted");
    }
    File.WriteAllText(tempSettings, "{\"port\":9000}");
    var externalFile = SettingsViewSaveCoordinator.FileFingerprint(tempSettings);
    Check(externalFile != observedFile, "external settings edit was not detected");
    bool persistedStale = false;
    Check(!SettingsViewSaveCoordinator.TryCommit(current, () => current,
        () => persistedStale = true,
        () => SettingsViewSaveCoordinator.FileFingerprint(tempSettings) == observedFile),
        "stale on-disk settings must reject a save");
    Check(!persistedStale, "stale on-disk settings unexpectedly persisted");
    // Even if the bytes have not changed since opening, values already
    // different from this process's in-memory settings must not be saved.
    Check(!SettingsViewSaveCoordinator.TryCommit(current, () => current,
        () => persistedStale = true,
        () => new SettingsViewSnapshot(9000, false, true).Equals(current)),
        "external settings values must reject a save");
    Check(SettingsViewSaveCoordinator.TryCommit(current, () => current,
        () => persistedStale = true,
        () => SettingsViewSaveCoordinator.FileFingerprint(tempSettings) == externalFile),
        "fresh on-disk state should allow saving");
    Check(persistedStale, "fresh on-disk state did not persist");
}
finally
{
    if (File.Exists(tempSettings)) File.Delete(tempSettings);
}

Console.WriteLine("Settings view snapshot tests passed");
