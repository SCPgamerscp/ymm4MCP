using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var complete = new Mp4Inspection(true, "", "isom", 120.2, 1920, 1080, true);
var pass = ExportAcceptance.Evaluate("out.mp4", complete, 120, 0.5, 1920, 1080, true);
Check(pass.Success && pass.Passed && pass.Issues.Length == 0, "valid export failed");

var incomplete = new Mp4Inspection(true, "", "isom", 90, 1280, 720, false);
var fail = ExportAcceptance.Evaluate("out.mp4", incomplete, 120, 1, 1920, 1080, true);
Check(!fail.Passed && fail.Issues.Length == 4, "missing expected criteria");
Check(fail.Issues.Any(i => i.Code == "AUDIO_STREAM_MISSING"), "audio issue missing");
var invalid = ExportAcceptance.Evaluate("out.mp4", new Mp4Inspection(false, "truncated"));
Check(!invalid.Success && !invalid.Passed && invalid.Issues[0].Code == "EXPORT_VERIFY_FAILED",
    "invalid MP4 should fail");
try { ExportAcceptance.Evaluate("out.mp4", complete, double.NaN); throw new Exception("NaN accepted"); }
catch (ArgumentException) { }
Console.WriteLine("Export acceptance tests passed");
