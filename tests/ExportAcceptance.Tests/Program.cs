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
var fpsInspection = complete with { AverageFps = 29.97 };
Check(ExportAcceptance.Evaluate("out.mp4", fpsInspection, expectedFps: 30).Passed,
    "fractional FPS within tolerance rejected");
var fpsFail = ExportAcceptance.Evaluate("out.mp4", fpsInspection, expectedFps: 60);
Check(!fpsFail.Passed && fpsFail.Issues[0].Code == "FPS_MISMATCH", "FPS mismatch not reported");
Check(!ExportAcceptance.Evaluate("out.mp4", complete, expectedFps: 30).Passed,
    "missing frame timing accepted as matching FPS");
try { ExportAcceptance.Evaluate("out.mp4", fpsInspection, expectedFps: double.NaN); throw new Exception("NaN FPS accepted"); }
catch (ArgumentException) { }
Console.WriteLine("Export acceptance tests passed");
