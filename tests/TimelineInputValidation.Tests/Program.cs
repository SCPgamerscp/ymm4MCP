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

var overlap = new List<(Item item, int frame, int length, int layer)>
{
    (first, 0, 10, 0), (last, 5, 5, 0)
};
edits = 0;
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

Console.WriteLine("Timeline input validation tests passed");

sealed class Item(int frame)
{
    public int Frame { get; set; } = frame;
}
