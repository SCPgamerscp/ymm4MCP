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
