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
Console.WriteLine("Settings view snapshot tests passed");
