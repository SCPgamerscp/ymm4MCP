namespace YMM4McpPlugin
{
    /// <summary>
    /// Captures the server settings observed when a tool window was opened or
    /// last saved, preventing stale views from silently overwriting newer edits.
    /// </summary>
    internal readonly record struct SettingsViewSnapshot(int Port, bool AutoStart, bool AllowAdvanced)
    {
        public bool Matches(int port, bool autoStart, bool allowAdvanced) =>
            Port == port && AutoStart == autoStart && AllowAdvanced == allowAdvanced;
    }
    /// <summary>
    /// Serializes the compare-and-save sequence across YMM4 tool windows.
    /// The persistence callback runs only when the observed settings still
    /// match the shared state; an exception never counts as a successful save.
    /// </summary>
    internal static class SettingsViewSaveCoordinator
    {
        private static readonly object Gate = new();

        internal static bool TryCommit(SettingsViewSnapshot observed,
            Func<SettingsViewSnapshot> readCurrent, Action persist)
        {
            lock (Gate)
            {
                if (!observed.Equals(readCurrent())) return false;
                persist();
                return true;
            }
        }
    }
}
