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
// Exercise the complete coordinate-selection path, not only IsAmbiguousMatch.
// A previous early break in LocateTimelineItem made the ambiguity guard dead code.
static (object? item, bool ambiguous) SelectAt(
    int frame, int layer, params (object? item, int frame, int layer)[] entries)
    => KeyframeSelectorSafety.FindUniqueCoordinateItem(
        entries, frame, layer, entry =>
        {
            var value = (ValueTuple<object?, int, int>)entry;
            return (value.Item1, value.Item2, value.Item3);
        });

var second = new object();
var unique = SelectAt(15, 4, (second, 14, 4), (first, 15, 4), (second, 15, 5));
Check(ReferenceEquals(unique.item, first) && !unique.ambiguous,
    "one matching item must be selected even if it is not the first timeline item");
var overlap = SelectAt(15, 4, (first, 15, 4), (second, 15, 4));
Check(overlap.item == null && overlap.ambiguous,
    "two distinct items at identical coordinates must be rejected");
var lateOverlap = SelectAt(15, 4, (first, 15, 4), (second, 1, 4),
    (first, 15, 4), (second, 15, 4));
Check(lateOverlap.item == null && lateOverlap.ambiguous,
    "a duplicate discovered late in enumeration must not be silently ignored");
var repeated = SelectAt(15, 4, (first, 15, 4), (first, 15, 4));
Check(ReferenceEquals(repeated.item, first) && !repeated.ambiguous,
    "repeated wrappers for the same item reference must not be ambiguous");
var empty = SelectAt(15, 4, (second, 16, 4), (null, 15, 4));
Check(empty.item == null && !empty.ambiguous,
    "no matching non-null item must remain not found");
var count = 0;
var inspected = KeyframeSelectorSafety.FindUniqueCoordinateItem(
    new object[] { first, second, first }, 15, 4, entry =>
    {
        count++;
        return (entry, 15, 4);
    });
Check(inspected.ambiguous && count == 2,
    "selection must inspect beyond the first match and stop after proving ambiguity");

// Verify the actual mutation result, not just whether a reflected method
// returned normally. Native YMM4 methods can silently ignore a request.
static bool Verified(string action, int at, double? requested,
    int[] frames, double[] values, double? fallback = null)
    => KeyframeMutationVerification.Confirm(action, at, requested, frames, values, fallback);

Check(KeyframeMutationVerification.IsInItemRange(0, 30), "first frame is in range");
Check(KeyframeMutationVerification.IsInItemRange(29, 30), "last visible frame is in range");
Check(!KeyframeMutationVerification.IsInItemRange(30, 30), "end-exclusive frame must be rejected");
Check(!KeyframeMutationVerification.IsInItemRange(-1, 30), "negative frame must be rejected");
Check(!KeyframeMutationVerification.IsInItemRange(0, 0), "unknown item duration must fail closed");
Check(Verified("set", 0, 1.25, Array.Empty<int>(), new[] { 1.25 }),
    "default-only animation must confirm a changed starting value");
Check(Verified("set", 0, 1.25, Array.Empty<int>(), Array.Empty<double>(), 1.25),
    "default-only host property without Values must remain supported");
Check(!Verified("set", 0, 1.25, Array.Empty<int>(), Array.Empty<double>(), 2.0),
    "a no-op native setter must not claim success");
Check(!Verified("set", 10, 1.25, Array.Empty<int>(), new[] { 1.25 }),
    "native add with no frame must fail readback");
Check(Verified("set", 10, 2.5, new[] { 10 }, new[] { 1.0, 2.5 }),
    "implicit start + inserted frame must confirm the correct value");
Check(Verified("set", 0, 1.0, new[] { 10 }, new[] { 1.0, 2.5 }),
    "implicit start frame zero must be readable");
Check(Verified("set", 10, 2.5, new[] { 0, 10 }, new[] { 1.0, 2.5 }),
    "explicit start layout must be readable");
Check(!Verified("set", 10, 2.5, new[] { 10 }, new[] { 1.0, 3.0 }),
    "a wrong value must fail even when the frame exists");
Check(!Verified("set", 10, 2.5, new[] { 10 }, new[] { 1.0 }),
    "partial native insertion without a new value must fail");
Check(!Verified("set", 10, 2.5, new[] { 10, 10 }, new[] { 1.0, 2.5 }),
    "duplicate frame indices must not be acknowledged");
Check(Verified("set", 10, 1.25, new[] { 10 }, new[] { 1.0, (double)(float)1.25 }),
    "float-backed animation values must be accepted");
Check(!Verified("set", 10, double.NaN, new[] { 10 }, new[] { 1.0, 1.25 }),
    "non-finite requested value must fail");
Check(KeyframeMutationVerification.HasObservableChange(
        new[] { 0 }, new[] { 1.0 }, new[] { 0, 10 }, new[] { 1.0 }),
    "partial frame-only insertion must invalidate optimistic revisions");
Check(KeyframeMutationVerification.HasObservableChange(
        new[] { 0 }, new[] { 1.0 }, new[] { 0 }, new[] { 2.0 }),
    "failed native setter changing a value must invalidate revisions");
Check(!KeyframeMutationVerification.HasObservableChange(
        new[] { 0 }, new[] { 1.0 }, new[] { 0 }, new[] { 1.0 }),
    "unchanged native failures must not spuriously advance revisions");
Check(Verified("remove", 10, null, new[] { 0 }, new[] { 1.0 }),
    "native removal must confirm frame absence");
Check(!Verified("remove", 10, null, new[] { 0, 10 }, new[] { 1.0, 2.5 }),
    "native no-op removal must fail");
Check(!Verified("remove", 10, null, new[] { 0 }, new[] { 1.0, 2.5, 3.0 }),
    "native remove that leaves orphaned values must fail");
Check(Verified("remove", 10, null, new[] { 0 }, new[] { 1.0, 2.5 }),
    "native remove may preserve an implicit starting value");
Check(Verified("clear", 0, null, Array.Empty<int>(), new[] { 1.0 }),
    "cleared frames may retain the base value");
Check(Verified("clear", 0, null, Array.Empty<int>(), Array.Empty<double>()),
    "cleared frames may remove all values");
Check(!Verified("clear", 0, null, Array.Empty<int>(), new[] { 1.0, 2.5 }),
    "native clear leaving orphaned values must fail");
Check(!Verified("clear", 0, null, new[] { 10 }, new[] { 1.0, 2.5 }),
    "native clear leaving frames must fail");

Console.WriteLine("Keyframe selector and mutation readback safety tests passed");
