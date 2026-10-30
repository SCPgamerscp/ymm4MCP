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
// The requested key can be correct even when the host corrupts other keys.
static bool Preserved(string action, int at, int[] beforeFrames, double[] beforeValues,
    int[] afterFrames, double[] afterValues)
    => KeyframeMutationVerification.PreservesUnrelatedKeys(
        action, at, beforeFrames, beforeValues, afterFrames, afterValues);
Check(Preserved("set", 10, Array.Empty<int>(), new[] { 1.0 },
        new[] { 10 }, new[] { 1.0, 2.5 }),
    "new key insertion must preserve implicit base value");
Check(Preserved("set", 10, new[] { 0, 10, 20 }, new[] { 1.0, 2.0, 3.0 },
        new[] { 0, 10, 20 }, new[] { 1.0, 2.5, 3.0 }),
    "editing a key may change only the requested value");
Check(Preserved("remove", 10, new[] { 0, 10, 20 }, new[] { 1.0, 2.0, 3.0 },
        new[] { 0, 20 }, new[] { 1.0, 3.0 }),
    "removal must preserve other explicit key values");
Check(Preserved("remove", 10, new[] { 10, 20 }, new[] { 1.0, 2.0, 3.0 },
        new[] { 20 }, new[] { 1.0, 3.0 }),
    "removal must preserve implicit start value");
Check(!Preserved("set", 10, new[] { 0, 10 }, new[] { 1.0, 2.0 },
        new[] { 0, 10 }, new[] { 99.0, 2.5 }),
    "native success must not overwrite a different key");
Check(!Preserved("remove", 10, new[] { 0, 10, 20 }, new[] { 1.0, 2.0, 3.0 },
        new[] { 0 }, new[] { 1.0 }),
    "native removal must not delete an unrelated key");
Check(!Preserved("set", 10, new[] { 0 }, new[] { 1.0 },
        new[] { 0, 10, 20 }, new[] { 1.0, 2.5, 9.0 }),
    "native insertion must not introduce unrelated keys");
Check(!Preserved("set", 10, new[] { 0 }, new[] { 1.0 },
        new[] { 0, 10 }, new[] { 1.0 }),
    "a frame-only partial insertion must fail preservation");
Check(!Preserved("set", 10, new[] { 0 }, new[] { 1.0 },
        new[] { 0, 10, 10 }, new[] { 1.0, 2.5, 2.5 }),
    "duplicate frame layout must fail preservation");
Check(!Preserved("set", 10, new[] { 0 }, new[] { 1.0 },
        new[] { 0, 10 }, new[] { 0.5, 1.0, 2.5 }),
    "implicit start combined with explicit zero must fail as ambiguous");
Check(Preserved("set", 10, new[] { 20 }, new[] { 1.0, 3.0 },
        new[] { 0, 10, 20 }, new[] { 1.0, 2.5, 3.0 }),
    "logical values may survive a change from implicit to explicit start layout");
Check(Preserved("clear", 0, new[] { 0, 10 }, new[] { 1.0, 2.5 },
        Array.Empty<int>(), Array.Empty<double>()),
    "clear intentionally removes all unrelated keys");
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

// The GET inventory must not report broken frame/value layouts as valid.
// Strict inventory fails closed; best-effort inventory skips that property.
Check(KeyframeMutationVerification.HasReadableLayout(
        Array.Empty<int>(), Array.Empty<double>()),
    "a default-only animation without a Values array is readable");
Check(KeyframeMutationVerification.HasReadableLayout(
        new[] { 10, 20 }, new[] { 1.0, 2.0, 3.0 }),
    "implicit start value with two explicit frames is readable");
Check(KeyframeMutationVerification.HasReadableLayout(
        new[] { 0, 10 }, new[] { 1.0, 2.0 }),
    "explicit frame zero with aligned values is readable");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 10, 20 }, new[] { 1.0 }),
    "truncated host Values must not pass inventory readback");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 10 }, new[] { 1.0, 2.0, 3.0 }),
    "orphaned host Values must not pass inventory readback");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 10, 10 }, new[] { 1.0, 2.0 }),
    "duplicate host Frames must not pass inventory readback");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 20, 10 }, new[] { 2.0, 1.0 }),
    "out-of-order host Frames must not be advertised as a valid animation");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 20, 10 }, new[] { 1.0, 2.0, 3.0 }),
    "implicit-start host layout must still require sorted Frames");
Check(!Verified("set", 10, 1.0, new[] { 20, 10 },
        new[] { 2.0, 1.0 }),
    "a native setter must not report success with reordered keyframes");
Check(!Verified("remove", 30, null, new[] { 20, 10 },
        new[] { 2.0, 1.0 }),
    "a native remover must not report success with reordered keyframes");
Check(!Verified("clear", 0, null, new[] { 10, 5 },
        new[] { 2.0, 1.0 }),
    "a native clear must not report success with malformed remaining frames");
Check(!Preserved("set", 10, new[] { 20, 10 }, new[] { 2.0, 1.0 },
        new[] { 10, 20 }, new[] { 1.0, 2.0 }),
    "a malformed pre-edit layout must never be accepted as preserved");
Check(!Preserved("set", 10, new[] { 10, 20 }, new[] { 1.0, 2.0 },
        new[] { 20, 10 }, new[] { 2.0, 1.0 }),
    "a reordered post-edit layout must not be accepted as preserved");

Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 0, 10 }, new[] { 1.0, 2.0, 3.0 }),
    "implicit start and explicit frame zero cannot coexist unambiguously");
Check(!KeyframeMutationVerification.HasReadableLayout(
        new[] { 10 }, new[] { double.NaN, 2.0 }),
    "non-finite host values must not pass inventory readback");

// Host reflection readback must never silently discard malformed entries.
// Previously ReadFrameList/ReadValueList skipped them, which could make a
// partially corrupted native animation look like a successful edit.
static void RejectObserved(Action read, string reason)
{
    try { read(); }
    catch (InvalidOperationException) { return; }
    throw new Exception("Unsafe keyframe readback was accepted: " + reason);
}
var observedFrames = KeyframeMutationVerification.ParseObservedFrames(new object[] { 0, 12 });
Check(observedFrames.SequenceEqual(new[] { 0, 12 }),
    "valid host frames must preserve their order and count");
var observedValues = KeyframeMutationVerification.ParseObservedValues(
    new object[] { new { Value = 1.25 }, new { Value = 2.5 } },
    entry => entry.GetType().GetProperty("Value")?.GetValue(entry));
Check(observedValues.SequenceEqual(new[] { 1.25, 2.5 }),
    "wrapped host values must preserve their order and count");
RejectObserved(() => KeyframeMutationVerification.ParseObservedFrames(
    new object?[] { 0, null, 12 }), "null frame must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedFrames(
    new object[] { 0, "not-a-frame", 12 }), "malformed frame must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedFrames(
    new object[] { 0, -1, 12 }), "negative frame must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedValues(
    new object?[] { 1.0, null, 2.0 }, entry => entry),
    "null value must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedValues(
    new object[] { 1.0, "invalid", 2.0 }, entry => entry),
    "malformed value must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedValues(
    new object[] { 1.0, double.NaN, 2.0 }, entry => entry),
    "NaN value must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedValues(
    new object[] { 1.0, double.PositiveInfinity, 2.0 }, entry => entry),
    "infinite value must not disappear");
RejectObserved(() => KeyframeMutationVerification.ParseObservedValues(
    new object[] { 1.0 }, entry => null),
    "unreadable wrapped value must not disappear");

Console.WriteLine("Keyframe selector and mutation readback safety tests passed");
