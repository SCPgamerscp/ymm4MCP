using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace YMM4McpPlugin
{
    internal static class EditCheckpointSafety
    {
        internal const int MaxBindings = 5000;

        internal static void ValidateBinding(string logicalId, string itemId, HashSet<string> seenLogicalIds)
        {
            static bool Invalid(string value) =>
                string.IsNullOrWhiteSpace(value) || value.Length > 512 ||
                value.Trim() != value || value.Any(char.IsControl);

            if (Invalid(logicalId) || Invalid(itemId))
                throw new ArgumentException("binding id and item_id must be nonempty, unambiguous identifiers");
            if (!seenLogicalIds.Add(logicalId))
                throw new ArgumentException("duplicate logical binding id: " + logicalId);
        }

        // Reconcile may extend a partially applied plan, or replace a bound item
        // that has been deleted. Never remap a binding while its old native item
        // still exists, and never persist a binding to a nonexistent item.
        internal static bool CompatibleBindingUpdate(
            IEnumerable<(string Id, string ItemId, string? Revision)> previous,
            IEnumerable<(string Id, string ItemId, string? Revision)> proposed,
            IReadOnlyCollection<string>? liveItemIds)
        {
            if (previous == null || proposed == null || liveItemIds == null)
                return false;
            var live = new HashSet<string>(liveItemIds, StringComparer.Ordinal);
            if (live.Count != liveItemIds.Count || live.Any(string.IsNullOrWhiteSpace))
                return false;
            var incoming = new Dictionary<string, (string ItemId, string? Revision)>(StringComparer.Ordinal);
            var claimedItems = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in proposed)
            {
                if (!live.Contains(item.ItemId) || !claimedItems.Add(item.ItemId)
                    || !incoming.TryAdd(item.Id, (item.ItemId, item.Revision)))
                    return false;
            }
            var oldIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in previous)
            {
                if (!oldIds.Add(item.Id)) return false;
                if (live.Contains(item.ItemId))
                {
                    if (!incoming.TryGetValue(item.Id, out var next)
                        || !string.Equals(next.ItemId, item.ItemId, StringComparison.Ordinal)
                        || !string.Equals(next.Revision, item.Revision, StringComparison.Ordinal))
                        return false;
                }
            }
            return true;
        }

        // Persisted edits.json is untrusted input, even if normal API writes were
        // validated. Never restore malformed or ambiguous idempotency bindings.
        internal static bool ValidStoredBindingRecord(
            string? key, string? planHash,
            IEnumerable<(string? Id, string? ItemId)>? items)
        {
            if (string.IsNullOrWhiteSpace(key) || key.Length > 128 || key.Trim() != key
                || planHash == null || planHash.Length is < 16 or > 128
                || items == null)
                return false;
            var logicalIds = new HashSet<string>(StringComparer.Ordinal);
            var nativeIds = new HashSet<string>(StringComparer.Ordinal);
            int count = 0;
            foreach (var (id, itemId) in items)
            {
                if (++count > MaxBindings || id == null || itemId == null
                    || !nativeIds.Add(itemId))
                    return false;
                try { ValidateBinding(id, itemId, logicalIds); }
                catch (ArgumentException) { return false; }
            }
            return true;
        }

        // A checkpoint is only safe to use when every native item can be identified
        // uniquely. An empty list is valid for a genuinely empty timeline.
        internal static void ValidateSnapshotItemIds(IEnumerable<string> itemIds)
        {
            if (itemIds == null) throw new ArgumentException("checkpoint item ids are missing");
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var id in itemIds)
            {
                if (string.IsNullOrWhiteSpace(id) || id.Trim() != id ||
                    id.Any(char.IsControl) || !seen.Add(id))
                    throw new ArgumentException("checkpoint item ids are missing, malformed, or duplicated");
            }
        }

        // Rollback deletes items. Never apply a checkpoint to an unknown,
        // unsaved, or different project even when timeline IDs happen to match.
        internal static bool SameProject(string? expected, string? current)
        {
            if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(current))
                return false;
            try
            {
                return string.Equals(Path.GetFullPath(expected), Path.GetFullPath(current),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
            {
                return false;
            }
        }

        // Checkpoint records are restored from JSON. Never allow a forged
        // BackupPath to delete files outside the private checkpoint directory.
        internal static bool IsOwnedBackupPath(string? path, string backupDirectory)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(backupDirectory))
                return false;
            try
            {
                string full = Path.GetFullPath(path);
                string parent = Path.GetDirectoryName(full) ?? "";
                string name = Path.GetFileName(full);
                return string.Equals(parent, Path.GetFullPath(backupDirectory),
                           StringComparison.OrdinalIgnoreCase) &&
                       name.StartsWith("cp_", StringComparison.Ordinal) &&
                       name.EndsWith(".ymmp", StringComparison.OrdinalIgnoreCase) &&
                       name.Length is >= 16 and <= 72 &&
                       name[3..^5].All(char.IsLetterOrDigit) &&
                       name[3..^5].All(ch => ch <= 127);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
            {
                return false;
            }
        }

        // A failed copy must never leave a file with the final .ymmp backup name.
        internal static void CopyBackupAtomically(string source, string destination)
        {
            string temp = destination + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.Copy(source, temp);
                File.Move(temp, destination);
            }
            finally
            {
                try { File.Delete(temp); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }
}
