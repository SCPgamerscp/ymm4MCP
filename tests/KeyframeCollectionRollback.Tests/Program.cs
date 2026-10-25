using System;
using System.Collections;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

static void CheckUnchanged(IList frames, IList values, string scenario)
{
    Check(frames.Count == 1 && Convert.ToInt32(frames[0]) == 10,
        scenario + ": original frames were corrupted");
    Check(values.Count == 2 && Convert.ToDouble(values[0]) == 1.0
        && Convert.ToDouble(values[1]) == 2.0,
        scenario + ": original values were corrupted");
}

var normal = new TestAnimation(new FaultyList { 10 }, new FaultyList { 1.0, 2.0 });
Check(KeyframeCollectionMutation.TryUpsert(normal, 20, 5.0),
    "valid insertion should succeed");
Check(normal.Frames.Count == 2 && Convert.ToInt32(normal.Frames[1]) == 20
    && normal.Values.Count == 3 && Convert.ToDouble(normal.Values[2]) == 5.0,
    "valid insertion produced the wrong keyframe");

// A value collection can mutate and then throw. It must not leave an orphan.
var valueThrows = new TestAnimation(new FaultyList { 10 },
    new FaultyList { 1.0, 2.0 } { ThrowAfterInsert = true });
Check(!KeyframeCollectionMutation.TryUpsert(valueThrows, 20, 5.0),
    "throwing value insertion should fail");
CheckUnchanged(valueThrows.Frames, valueThrows.Values, "value throws after insertion");

// A frame collection can mutate and then throw. Restore both collections.
var frameThrows = new TestAnimation(
    new FaultyList { 10 } { ThrowAfterInsert = true }, new FaultyList { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(frameThrows, 20, 5.0),
    "throwing frame insertion should fail");
CheckUnchanged(frameThrows.Frames, frameThrows.Values, "frame throws after insertion");

// A non-conforming IList may silently ignore Insert. Rollback must not delete
// a pre-existing frame or value when the collection count did not increase.
var frameNoOp = new TestAnimation(
    new FaultyList { 10 } { IgnoreInsert = true }, new FaultyList { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(frameNoOp, 20, 5.0),
    "ignored frame insertion should fail");
CheckUnchanged(frameNoOp.Frames, frameNoOp.Values, "frame insert ignored");

var valueNoOp = new TestAnimation(new FaultyList { 10 },
    new FaultyList { 1.0, 2.0 } { IgnoreInsert = true });
Check(!KeyframeCollectionMutation.TryUpsert(valueNoOp, 20, 5.0),
    "ignored value insertion should fail");
CheckUnchanged(valueNoOp.Frames, valueNoOp.Values, "value insert ignored");

// Immutable-style list properties can throw on re-read after initial discovery.
var flaky = new FlakyImmutableAnimation();
Check(!KeyframeCollectionMutation.TryUpsert(flaky, 20, 5.0),
    "throwing immutable property getter escaped the fallback");
Console.WriteLine("Keyframe collection rollback tests passed");

internal sealed class TestAnimation
{
    public TestAnimation(IList frames, IList values) { Frames = frames; Values = values; }
    public IList Frames { get; set; }
    public IList Values { get; set; }
}

internal sealed class FaultyList : ArrayList
{
    public bool ThrowAfterInsert { get; set; }
    public bool IgnoreInsert { get; set; }
    public override void Insert(int index, object? value)
    {
        if (IgnoreInsert) return;
        base.Insert(index, value);
        if (ThrowAfterInsert) throw new InvalidOperationException("insert failed after mutation");
    }
}

internal sealed class FlakyImmutableAnimation
{
    private IList frames = ArrayList.FixedSize(new ArrayList { 10 });
    private IList values = ArrayList.FixedSize(new ArrayList { 1.0, 2.0 });
    private int frameReads;
    public IList Frames
    {
        get
        {
            frameReads++;
            if (frameReads > 1) throw new InvalidOperationException("window closed");
            return frames;
        }
        set => frames = value;
    }
    public IList Values { get => values; set => values = value; }
}
