using System;
using System.IO;
using System.Text.Json;

namespace YMM4McpPlugin
{
    internal sealed record ExportProjectIdentity(Guid SessionId, Guid TimelineId,
        string? ProjectPath, bool IsSaved)
    {
        internal bool IsKnown =>
            SessionId != Guid.Empty && TimelineId != Guid.Empty &&
            (ProjectPath == null || Path.IsPathFullyQualified(ProjectPath));

        internal bool Matches(ExportProjectIdentity? active)
        {
            if (active == null || !IsKnown || !active.IsKnown) return false;
            if (SessionId == active.SessionId)
                return TimelineId == active.TimelineId &&
                    ((ProjectPath == null && active.ProjectPath == null) ||
                        EditCheckpointSafety.SameProject(ProjectPath, active.ProjectPath));
            // Runtime identities do not survive YMM4 restart. An untitled or
            // unsaved source cannot be reconstructed from its saved path.
            return IsSaved && EditCheckpointSafety.IsStableSavedProject(
                ProjectPath, active.ProjectPath, active.IsSaved);
        }

        internal static ExportProjectIdentity? ReadSnapshot(JsonElement? snapshot,
            JsonSerializerOptions? options = null)
        {
            if (!snapshot.HasValue || snapshot.Value.ValueKind != JsonValueKind.Object) return null;
            try
            {
                var source = snapshot.Value.Deserialize<ExportProjectIdentity>(options);
                return source is { IsKnown: true } ? source : null;
            }
            catch (JsonException) { return null; }
        }
    }
}
