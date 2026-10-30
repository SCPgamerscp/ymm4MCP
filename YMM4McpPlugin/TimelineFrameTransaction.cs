using System;
using System.Collections.Generic;

namespace YMM4McpPlugin;

// Execute already-preflighted frame moves on the YMM4 UI dispatcher. If any
// setter throws or lies about the result, attempt to restore *every* touched
// item, including the setter that threw after mutating its target. This is a
// frame-only compensation, not a full host/project transaction.
internal static class TimelineFrameTransaction
{
    internal sealed record Result<T>(
        bool Success, bool RollbackVerified, bool OutcomeUnknown,
        string ErrorCode, IReadOnlyList<T> Touched) where T : class;

    internal static Result<T> Apply<T>(
        IReadOnlyList<(T item, int original, int target)> changes,
        Func<T, int> read, Action<T, int> write) where T : class
    {
        ArgumentNullException.ThrowIfNull(changes);
        ArgumentNullException.ThrowIfNull(read);
        ArgumentNullException.ThrowIfNull(write);
        var seen = new HashSet<T>(ReferenceEqualityComparer.Instance);
        // Validate all identities and current host values before the first
        // write; a stale later item must not partially shift earlier items.
        foreach (var change in changes)
        {
            if (change.item == null || change.original < 0 || change.target < 0 ||
                !seen.Add(change.item))
                return new(false, false, false, "FRAME_BATCH_PREFLIGHT_INVALID", Array.Empty<T>());
            try
            {
                if (read(change.item) != change.original)
                    return new(false, false, false, "FRAME_BATCH_STALE", Array.Empty<T>());
            }
            catch (Exception)
            {
                return new(false, false, false, "FRAME_BATCH_PREFLIGHT_UNREADABLE", Array.Empty<T>());
            }
        }

        var touched = new List<T>();
        try
        {
            foreach (var change in changes)
            {
                if (change.original == change.target) continue;
                // A setter may change the host model and then throw. Record
                // its item BEFORE calling it so compensation includes it.
                touched.Add(change.item);
                write(change.item, change.target);
                if (read(change.item) != change.target)
                    throw new InvalidOperationException("Frame setter did not retain the requested value");
            }
            return new(true, false, false, "", touched);
        }
        catch (Exception)
        {
            bool restored = true;
            // Compensate in reverse order; never stop after one rollback error.
            for (int i = touched.Count - 1; i >= 0; --i)
            {
                T item = touched[i];
                int original = 0;
                foreach (var change in changes)
                    if (ReferenceEquals(change.item, item))
                    {
                        original = change.original;
                        break;
                    }
                try
                {
                    write(item, original);
                    if (read(item) != original) restored = false;
                }
                catch (Exception)
                {
                    restored = false;
                }
            }
            // Frame readback cannot prove that the host restored every
            // property, undo entry, animation or side effect. Never label a
            // failed multi-item mutation as atomic success.
            return new(false, restored, touched.Count > 0,
                restored ? "FRAME_BATCH_COMPENSATED" : "FRAME_BATCH_PARTIAL",
                touched);
        }
    }
}
