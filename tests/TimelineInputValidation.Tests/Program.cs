using System.Text.Json;
using YMM4McpPlugin;

static Dictionary<string, JsonElement> Body(string json) =>
    JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json)!;

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new Exception($"Expected {expected}, got {actual}");
}

static void Invalid(Action action)
{
    try { action(); }
    catch (ArgumentException) { return; }
    throw new Exception("Expected ArgumentException");
}

Equal(7, TimelineInputValidation.GetInt(Body("{}"), "frame", 7));
Equal(-1, TimelineInputValidation.GetInt(Body("{\"frame\":-1}"), "frame", 7));
Equal(int.MaxValue, TimelineInputValidation.GetInt(Body("{\"frame\":2147483647}"), "frame", 7));
foreach (string bad in new[] { "null", "true", "\"3\"", "3.5", "2147483648", "-2147483649" })
    Invalid(() => TimelineInputValidation.GetInt(Body($"{{\"frame\":{bad}}}"), "frame", 7));

Equal<int[]?>(null, TimelineInputValidation.GetLayers(Body("{}")));
Equal(0, TimelineInputValidation.GetLayers(Body("{\"layers\":[]}"))!.Length);
Equal(2, TimelineInputValidation.GetLayers(Body("{\"layers\":[0,2147483647]}"))!.Length);
foreach (string bad in new[] { "null", "3", "\"1\"", "[-1]", "[true]", "[1.2]", "[2147483648]" })
    Invalid(() => TimelineInputValidation.GetLayers(Body($"{{\"layers\":{bad}}}")));
Invalid(() => TimelineInputValidation.GetLayers(Body("{\"layers\":[" + string.Join(',', Enumerable.Repeat("0", 1001)) + "]}")));

TimelineInputValidation.RequireCoordinates(Body("{\"frame\":0,\"layer\":0}"));
foreach (string bad in new[] { "{}", "{\"frame\":0}", "{\"layer\":0}",
                            "{\"frame\":-1,\"layer\":0}" })
    Invalid(() => TimelineInputValidation.RequireCoordinates(Body(bad)));
TimelineInputValidation.RequireItemTarget(Body("{\"item_id\":\"native:a\"}"));
TimelineInputValidation.RequireItemTarget(Body("{\"frame\":1,\"layer\":2}"));
foreach (string bad in new[] { "{}", "{\"item_id\":\" \"}", "{\"item_id\":3}",
                            "{\"frame\":0}" })
    Invalid(() => TimelineInputValidation.RequireItemTarget(Body(bad)));
TimelineInputValidation.RequireSelectionTarget(Body("{\"frame\":10}"));
TimelineInputValidation.RequireSelectionTarget(Body("{\"layer\":2}"));
TimelineInputValidation.RequireSelectionTarget(Body("{\"clear\":true}"));
Invalid(() => TimelineInputValidation.RequireSelectionTarget(Body("{}")));
TimelineInputValidation.RequireItemTarget(Body("{\"item_id\":\"native:日本語 with space\"}"));
TimelineInputValidation.RequireSelectionTarget(Body("{\"item_id\":\"native:日本語 with space\"}"));
foreach (string escape in new[] { "\\u0000", "\\u0009", "\\u000a", "\\u000d", "\\u001f", "\\u007f" })
{
    string invalidId = "{\"item_id\":\"native:abc" + escape + "\"}";
    Invalid(() => TimelineInputValidation.RequireItemTarget(Body(invalidId)));
    Invalid(() => TimelineInputValidation.RequireSelectionTarget(Body(invalidId)));
}
Invalid(() => TimelineInputValidation.RequireSelectionTarget(Body("{\"item_id\":\" \"}")));
Invalid(() => TimelineInputValidation.RequireSelectionTarget(Body("{\"item_id\":42}")));
Invalid(() => TimelineInputValidation.RequireSelectionTarget(Body("{\"frame\":-1}")));

Invalid(() => TimelineInputValidation.AtLeast(-2, -1, "frame"));
Equal(-1, TimelineInputValidation.AtLeast(-1, -1, "frame"));
Invalid(() => TimelineInputValidation.CheckPlacement(-1, 1));
Invalid(() => TimelineInputValidation.CheckPlacement(0, 0));
Invalid(() => TimelineInputValidation.CheckPlacement(int.MaxValue, 1));
Equal(int.MaxValue - 1, TimelineInputValidation.CheckPlacement(int.MaxValue - 1L, 1));

var first = new Item(10);
var last = new Item(100);
var shift = new List<(Item item, int frame, int length, int layer)> { (first, 10, 5, 0), (last, 100, 5, 1) };
int edits = 0;
void Set(Item item, int frame) { edits++; item.Frame = frame; }
Invalid(() => TimelineInputValidation.ApplyShift(shift, -101, (item, _, frame, _) => Set(item, frame)));
Equal(0, edits);
Equal(10, first.Frame);
Invalid(() => TimelineInputValidation.ApplyShift(
    new List<(Item, int, int, int)> { (first, 10, 5, 0), (last, int.MaxValue - 5, 5, 1) },
    1, (item, _, frame, _) => Set(item, frame)));
Equal(0, edits);
TimelineInputValidation.ApplyShift(shift, -5, (item, _, frame, _) => Set(item, frame));
Equal(2, edits);
Equal(5, first.Frame);
Equal(95, last.Frame);

// A zero-delta shift must not write item properties or create spurious undo entries.
edits = 0;
TimelineInputValidation.ApplyShift(shift, 0, (item, _, frame, _) => Set(item, frame));
Equal(0, edits);

// Validate the original source as well as the shifted destination before any write.
Invalid(() => TimelineInputValidation.ApplyShift(
    new List<(Item, int, int, int)> { (first, -1, 5, 0) },
    2, (item, _, frame, _) => Set(item, frame)));
Invalid(() => TimelineInputValidation.ApplyShift(
    new List<(Item, int, int, int)> { (first, 5, 5, -1) },
    2, (item, _, frame, _) => Set(item, frame)));
Invalid(() => TimelineInputValidation.ApplyShift(
    new List<(Item, int, int, int)> { (first, 5, 0, 0) },
    2, (item, _, frame, _) => Set(item, frame)));
Equal(0, edits);
Equal(5, first.Frame);

var overlap = new List<(Item item, int frame, int length, int layer)>
{
    (first, 0, 10, 0), (last, 5, 5, 0)
};
edits = 0;
// A malformed layer must fail before overlap resolution moves any other item.
Invalid(() => TimelineInputValidation.ApplyResolve(
    new List<(Item, int, int, int)> { (first, 0, 10, 0), (last, 5, 5, -1) },
    0, (item, _, frame, _, _) => Set(item, frame)));
Equal(0, edits);
var overflow = new Item(int.MaxValue - 3);
var invalidOverlap = new List<(Item item, int frame, int length, int layer)>(overlap)
{
    (overflow, int.MaxValue - 3, 4, 0)
};
Invalid(() => TimelineInputValidation.ApplyResolve(invalidOverlap, 0,
    (item, _, frame, _, _) => Set(item, frame)));
Equal(0, edits);
TimelineInputValidation.ApplyResolve(overlap, 0, (item, _, frame, _, _) => Set(item, frame));
Equal(1, edits);
Equal(10, last.Frame);

// Batch shifts must never report success after a later setter fails.
var batchA = new FrameItem(10);
var batchB = new FrameItem(20);
var goodBatch = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 10, 5), (batchB, 20, 15) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal(true, goodBatch.Success);
Equal(2, goodBatch.Touched.Count);
Equal(5, batchA.Frame);
Equal(15, batchB.Frame);

// A stale second item must reject the entire batch before any setter.
var stale = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 5, 3), (batchB, 20, 10) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal(false, stale.Success);
Equal("FRAME_BATCH_STALE", stale.ErrorCode);
Equal(0, stale.Touched.Count);
Equal(5, batchA.Frame);

// Duplicate host references cannot be mutated twice under different plans.
var duplicate = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 5, 2), (batchA, 5, 1) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal("FRAME_BATCH_PREFLIGHT_INVALID", duplicate.ErrorCode);
Equal(5, batchA.Frame);

// A setter that writes and then throws must itself be compensated, as well
// as all earlier successful setters.
batchB.ThrowAfter.Add(30);
var compensated = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 5, 4), (batchB, 15, 30) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal(false, compensated.Success);
Equal("FRAME_BATCH_COMPENSATED", compensated.ErrorCode);
Equal(true, compensated.RollbackVerified);
Equal(true, compensated.OutcomeUnknown);
Equal(2, compensated.Touched.Count);
Equal(5, batchA.Frame);
Equal(15, batchB.Frame);

// If even one compensating setter fails, report an unknown partial outcome.
batchB.ThrowAfter.Clear();
batchB.ThrowAfter.Add(15);
batchA.ThrowAfter.Add(4);
var partial = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 5, 4), (batchB, 15, 30) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal(false, partial.Success);
Equal("FRAME_BATCH_PARTIAL", partial.ErrorCode);
Equal(false, partial.RollbackVerified);
Equal(true, partial.OutcomeUnknown);
batchB.ThrowAfter.Clear();
batchA.ThrowAfter.Clear();

// Silent no-op setters are failures, not successful frame shifts.
batchB.Ignore.Add(30);
var ignored = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)>
    { (batchA, 5, 4), (batchB, 15, 30) },
    item => item.Frame, (item, frame) => item.Set(frame));
Equal(false, ignored.Success);
Equal(true, ignored.RollbackVerified);
Equal(5, batchA.Frame);
Equal(15, batchB.Frame);
batchB.Ignore.Clear();

// Zero-delta batches must not invoke a setter or mark any item touched.
var noop = TimelineFrameTransaction.Apply(
    new List<(FrameItem item, int original, int target)> { (batchA, 5, 5) },
    item => item.Frame, (_, _) => throw new Exception("unexpected write"));
Equal(true, noop.Success);
Equal(0, noop.Touched.Count);

string temporary = Path.Combine(Path.GetTempPath(), "ymm4mcp-backup-test-" + Guid.NewGuid().ToString("N"));
try
{
    Directory.CreateDirectory(temporary);
    string source = Path.Combine(temporary, "project.ymmp");
    string backups = Path.Combine(temporary, "backups");
    Equal<string?>(null, ProjectBackup.BeforeOverwrite(source, backups));
    File.WriteAllText(source, "first version");
    string firstBackup = ProjectBackup.BeforeOverwrite(source, backups)!;
    File.WriteAllText(source, "second version");
    string secondBackup = ProjectBackup.BeforeOverwrite(source, backups)!;
    Equal("first version", File.ReadAllText(firstBackup));
    Equal("second version", File.ReadAllText(secondBackup));
    Equal("second version", File.ReadAllText(source));
    Equal(false, firstBackup == secondBackup);
    Equal(0, Directory.GetFiles(backups, "*.partial").Length);
    // On Windows an exclusive source lock must not leave a published backup.
    if (OperatingSystem.IsWindows())
    {
        int countBeforeFailure = Directory.GetFiles(backups, "*.ymmp").Length;
        using (var locked = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            bool failed = false;
            try { ProjectBackup.BeforeOverwrite(source, backups); }
            catch (IOException) { failed = true; }
            Equal(true, failed);
        }
        Equal(countBeforeFailure, Directory.GetFiles(backups, "*.ymmp").Length);
        Equal(0, Directory.GetFiles(backups, "*.partial").Length);
    }
}
finally { if (Directory.Exists(temporary)) Directory.Delete(temporary, recursive: true); }

Console.WriteLine("Timeline input validation tests passed");

sealed class FrameItem(int frame)
{
    public int Frame { get; private set; } = frame;
    public HashSet<int> ThrowAfter { get; } = new();
    public HashSet<int> Ignore { get; } = new();
    public void Set(int value)
    {
        if (!Ignore.Contains(value)) Frame = value;
        if (ThrowAfter.Contains(value)) throw new InvalidOperationException("Injected setter failure");
    }
}

sealed class Item(int frame)
{
    public int Frame { get; set; } = frame;
}
