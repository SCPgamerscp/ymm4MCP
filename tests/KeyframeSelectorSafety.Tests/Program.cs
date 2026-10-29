using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
static void Reject(string? id, string? frame, string? layer)
{
    try { KeyframeSelectorSafety.Parse(id, frame, layer); }
    catch (ArgumentException) { return; }
    throw new Exception($"Unexpectedly accepted selector id={id} frame={frame} layer={layer}");
}

var byId = KeyframeSelectorSafety.Parse("native:123", null, null);
Check(byId.itemId == "native:123" && byId.frame == -1 && byId.layer == -1,
    "native item_id must remain authoritative");
Check(KeyframeSelectorSafety.Parse("native:123", "bad", "bad").itemId == "native:123",
    "legacy coordinate hints must not change item_id selection");
var byCoords = KeyframeSelectorSafety.Parse(null, "123", "7");
Check(byCoords.itemId == "" && byCoords.frame == 123 && byCoords.layer == 7,
    "complete coordinates must be parsed exactly");
Check(KeyframeSelectorSafety.Parse(null, "0", "0").frame == 0,
    "explicit origin remains valid");
Reject(null, null, null);
Reject(null, "1", null);
Reject(null, null, "2");
Reject(null, "not-a-number", "0");
Reject(null, "0", "not-a-number");
Reject(null, "-1", "0");
Reject(null, "0", "-1");
Reject(null, "2147483648", "0");
Reject(null, "0", "2147483648");
Reject(null, " 1", "0");
Reject(null, "1 ", "0");
Reject("  ", null, null);
Reject("native:\u202e123", null, null);
Reject(new string('x', 257), null, null);
var first = new object();
Check(!KeyframeSelectorSafety.IsAmbiguousMatch(null, first),
    "first match is not ambiguous");
Check(!KeyframeSelectorSafety.IsAmbiguousMatch(first, first),
    "repeated reference is not a distinct timeline item");
Check(KeyframeSelectorSafety.IsAmbiguousMatch(first, new object()),
    "distinct timeline items at the same coordinates must be ambiguous");
Check(KeyframeSelectorSafety.AbsoluteFrame(int.MaxValue - 1, 100) == (long)int.MaxValue + 99,
    "absolute frame arithmetic must not overflow int32");
Console.WriteLine("Keyframe selector safety tests passed");
