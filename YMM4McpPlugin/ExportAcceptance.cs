using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    internal sealed record ExportQaIssue(string Code, string Severity, string Message,
        object? Expected = null, object? Actual = null);
    internal sealed record ExportQaReport(bool Success, bool Passed, string OutputPath,
        string? Brand, double? DurationSeconds, int? Width, int? Height, bool HasAudio,
        ExportQaIssue[] Issues);

    internal static class ExportAcceptance
    {
        public static ExportQaReport Evaluate(string path, Mp4Inspection inspection,
            double? expectedDuration = null, double tolerance = 1,
            int? expectedWidth = null, int? expectedHeight = null, bool requireAudio = false)
        {
            if (expectedDuration is double duration && (!double.IsFinite(duration) || duration <= 0))
                throw new ArgumentException("expected_duration_seconds must be positive and finite");
            if (!double.IsFinite(tolerance) || tolerance < 0 || tolerance > 60)
                throw new ArgumentException("duration_tolerance_seconds must be in 0..60");
            if (expectedWidth is < 1 or > 16384 || expectedHeight is < 1 or > 16384)
                throw new ArgumentException("expected_width/height must be in 1..16384");
            var issues = new List<ExportQaIssue>();
            if (!inspection.Verified)
                issues.Add(new("EXPORT_VERIFY_FAILED", "error", inspection.Error));
            else
            {
                if (expectedDuration.HasValue && Math.Abs(inspection.DurationSeconds!.Value - expectedDuration.Value) > tolerance)
                    issues.Add(new("DURATION_MISMATCH", "error", "動画の尺が期待値から外れています",
                        expectedDuration, inspection.DurationSeconds));
                if (expectedWidth.HasValue && inspection.Width != expectedWidth)
                    issues.Add(new("WIDTH_MISMATCH", "error", "出力幅が期待値と異なります",
                        expectedWidth, inspection.Width));
                if (expectedHeight.HasValue && inspection.Height != expectedHeight)
                    issues.Add(new("HEIGHT_MISMATCH", "error", "出力高さが期待値と異なります",
                        expectedHeight, inspection.Height));
                if (requireAudio && !inspection.HasAudio)
                    issues.Add(new("AUDIO_STREAM_MISSING", "error", "音声トラックがありません", true, false));
            }
            return new(inspection.Verified, issues.Count == 0, path, inspection.Brand,
                inspection.DurationSeconds, inspection.Width, inspection.Height, inspection.HasAudio, issues.ToArray());
        }
    }
}
