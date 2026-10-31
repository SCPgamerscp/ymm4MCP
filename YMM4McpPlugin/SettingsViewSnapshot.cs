using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;

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
    /// Serializes compare-and-save across windows and YMM4 processes.
    /// The persisted file is checked while holding the named mutex so a
    /// second updated YMM4 process cannot overwrite a newer settings save.
    /// </summary>
    internal static class SettingsViewSaveCoordinator
    {
        internal const string UnreadableFingerprint = "<unreadable>";

        private static readonly object Gate = new();
        private static readonly Mutex ProcessGate = new(false, @"Local\YMM4MCP.SettingsSave");

        internal static string FileFingerprint(string path)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
                return Convert.ToHexString(SHA256.HashData(stream));
            }
            catch (FileNotFoundException) { return "<missing>"; }
            catch (DirectoryNotFoundException) { return "<missing>"; }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The file exists but its state cannot be verified. Never
                // mistake this for a missing file or a matching fingerprint.
                return UnreadableFingerprint;
            }
        }

        internal static bool SameReadableFingerprint(string observed, string current) =>
            observed != UnreadableFingerprint &&
            current != UnreadableFingerprint &&
            string.Equals(observed, current, StringComparison.Ordinal);

        internal static bool TryCommit(SettingsViewSnapshot observed,
            Func<SettingsViewSnapshot> readCurrent, Action persist,
            Func<bool>? diskUnchanged = null)
        {
            bool acquired = false;
            try
            {
                try { acquired = ProcessGate.WaitOne(TimeSpan.FromSeconds(10)); }
                catch (AbandonedMutexException) { acquired = true; }
                if (!acquired)
                    throw new TimeoutException("別のYMM4プロセスが設定を保存中です");
                lock (Gate)
                {
                    if (!observed.Equals(readCurrent()) ||
                        (diskUnchanged != null && !diskUnchanged()))
                        return false;
                    persist();
                    return true;
                }
            }
            finally
            {
                if (acquired) ProcessGate.ReleaseMutex();
            }
        }
    }
}