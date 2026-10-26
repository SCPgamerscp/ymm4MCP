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
