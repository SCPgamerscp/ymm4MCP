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
}
