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
Console.WriteLine("Settings view snapshot tests passed");
