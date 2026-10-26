using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var complete = new Mp4Inspection(true, "", "isom", 120.2, 1920, 1080, true,
    AudioSampleRate: 48000, AudioChannels: 2);
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
Check(ExportAcceptance.Evaluate("out.mp4", complete, expectedAudioSampleRate: 48000,
    expectedAudioChannels: 2).Passed, "matching audio format rejected");
var badAudio = ExportAcceptance.Evaluate("out.mp4", complete, expectedAudioSampleRate: 44100,
    expectedAudioChannels: 1);
Check(badAudio.Issues.Any(i => i.Code == "AUDIO_SAMPLE_RATE_MISMATCH") &&
      badAudio.Issues.Any(i => i.Code == "AUDIO_CHANNELS_MISMATCH"), "audio mismatches missing");
Check(!ExportAcceptance.Evaluate("out.mp4", complete with { AudioSampleRate = null },
    expectedAudioSampleRate: 48000).Passed, "unknown audio rate accepted");
try { ExportAcceptance.Evaluate("out.mp4", complete, expectedAudioChannels: 9); throw new Exception("invalid channels accepted"); }
catch (ArgumentException) { }
try { ExportAcceptance.Evaluate("out.mp4", fpsInspection, expectedFps: double.NaN); throw new Exception("NaN FPS accepted"); }
catch (ArgumentException) { }
// A verified MP4 may still lack usable duration or frame-timing metadata.
var noDuration = complete with { DurationSeconds = null };
var missingDuration = ExportAcceptance.Evaluate("out.mp4", noDuration, expectedDuration: 120);
Check(!missingDuration.Passed && missingDuration.Issues.Any(i => i.Code == "DURATION_MISMATCH"),
    "missing duration must report a mismatch instead of throwing");
Check(ExportAcceptance.Evaluate("out.mp4", noDuration).Passed,
    "duration metadata should remain optional without an expected duration");
foreach (var nonFinite in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity })
{
    var invalidDuration = ExportAcceptance.Evaluate("out.mp4",
        complete with { DurationSeconds = nonFinite }, expectedDuration: 120);
    Check(!invalidDuration.Passed && invalidDuration.Issues.Any(i => i.Code == "DURATION_MISMATCH"),
        "non-finite duration must fail the requested duration check");
    var invalidFps = ExportAcceptance.Evaluate("out.mp4",
        complete with { AverageFps = nonFinite }, expectedFps: 30);
    Check(!invalidFps.Passed && invalidFps.Issues.Any(i => i.Code == "FPS_MISMATCH"),
        "non-finite FPS must fail the requested frame-rate check");
}
// Audio-format expectations imply that an audio stream must exist, even if
// requireAudio was not explicitly set.
var phantomAudio = complete with { HasAudio = false };
var missingAudioStream = ExportAcceptance.Evaluate("out.mp4", phantomAudio,
    expectedAudioSampleRate: 48000, expectedAudioChannels: 2);
Check(!missingAudioStream.Passed && missingAudioStream.Issues.Any(i => i.Code == "AUDIO_STREAM_MISSING"),
    "audio metadata without an audio stream must not pass");

foreach (var bad in new[] { double.NaN, double.PositiveInfinity, -1.0, 0.0 })
{
    Check(!ExportAcceptance.Evaluate("out.mp4", complete with { DurationSeconds = bad }).Passed,
        "invalid actual duration must fail even without an expected duration");
    Check(!ExportAcceptance.Evaluate("out.mp4", complete with { AverageFps = bad }).Passed,
        "invalid actual FPS must fail even without an expected FPS");
}
Check(!ExportAcceptance.Evaluate("out.mp4", complete with { Width = -1 }).Passed,
    "negative actual resolution must fail");
Check(!ExportAcceptance.Evaluate("out.mp4", complete with { AudioChannels = 99 }).Passed,
    "impossible audio channels must fail");
Check(ExportAcceptance.Evaluate("out.mp4", complete with { AverageFps = null }).Passed,
    "unknown optional FPS metadata must remain allowed");
// Verified status must not mask contradictory or malformed container metadata.
foreach (var brand in new string?[] { null, "", "    ", "too-long", "is\\u0001" })
{
    var result = ExportAcceptance.Evaluate("out.mp4", complete with { Brand = brand });
    Check(!result.Passed && result.Issues.Any(i => i.Code == "INVALID_BRAND_METADATA"),
        "invalid verified MP4 brand passed QA");
}
Check(!ExportAcceptance.Evaluate("out.mp4", complete with { Width = 20000 }).Passed,
    "out-of-range verified width passed QA");
Check(!ExportAcceptance.Evaluate("out.mp4", complete with { Height = 20000 }).Passed,
    "out-of-range verified height passed QA");
var inconsistentAudio = ExportAcceptance.Evaluate("out.mp4", complete with { HasAudio = false });
Check(!inconsistentAudio.Passed &&
      inconsistentAudio.Issues.Any(i => i.Code == "INCONSISTENT_AUDIO_METADATA"),
    "audio format without an audio track passed QA");
Check(ExportAcceptance.Evaluate("out.mp4", incomplete).Passed,
    "valid silent export should remain accepted without an audio requirement");
Console.WriteLine("Export acceptance tests passed");
