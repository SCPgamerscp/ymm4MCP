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
var throwingValues = new FaultyList { 1.0, 2.0 };
throwingValues.ThrowAfterInsert = true;
var valueThrows = new TestAnimation(new FaultyList { 10 }, throwingValues);
Check(!KeyframeCollectionMutation.TryUpsert(valueThrows, 20, 5.0),
    "throwing value insertion should fail");
CheckUnchanged(valueThrows.Frames, valueThrows.Values, "value throws after insertion");

// A frame collection can mutate and then throw. Restore both collections.
var throwingFrames = new FaultyList { 10 };
throwingFrames.ThrowAfterInsert = true;
var frameThrows = new TestAnimation(throwingFrames, new FaultyList { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(frameThrows, 20, 5.0),
    "throwing frame insertion should fail");
CheckUnchanged(frameThrows.Frames, frameThrows.Values, "frame throws after insertion");

// A non-conforming IList may silently ignore Insert. Rollback must not delete
// a pre-existing frame or value when the collection count did not increase.
var ignoredFrames = new FaultyList { 10 };
ignoredFrames.IgnoreInsert = true;
var frameNoOp = new TestAnimation(ignoredFrames, new FaultyList { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(frameNoOp, 20, 5.0),
    "ignored frame insertion should fail");
CheckUnchanged(frameNoOp.Frames, frameNoOp.Values, "frame insert ignored");

var ignoredValues = new FaultyList { 1.0, 2.0 };
ignoredValues.IgnoreInsert = true;
var valueNoOp = new TestAnimation(new FaultyList { 10 }, ignoredValues);
Check(!KeyframeCollectionMutation.TryUpsert(valueNoOp, 20, 5.0),
    "ignored value insertion should fail");
CheckUnchanged(valueNoOp.Frames, valueNoOp.Values, "value insert ignored");

// A broken IList may insert at index zero instead of the requested index.
// The operation must fail, and rollback must remove the actual new element,
// not the pre-existing element at the requested index.
var misplacedValues = new FaultyList { 1.0, 2.0 };
misplacedValues.InsertAtFront = true;
var wrongValueIndex = new TestAnimation(new FaultyList { 10 }, misplacedValues);
Check(!KeyframeCollectionMutation.TryUpsert(wrongValueIndex, 20, 5.0),
    "misplaced value insertion was reported successful");
CheckUnchanged(wrongValueIndex.Frames, wrongValueIndex.Values, "value inserted at wrong index");

var misplacedFrames = new FaultyList { 10 };
misplacedFrames.InsertAtFront = true;
var wrongFrameIndex = new TestAnimation(misplacedFrames, new FaultyList { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(wrongFrameIndex, 20, 5.0),
    "misplaced frame insertion was reported successful");
CheckUnchanged(wrongFrameIndex.Frames, wrongFrameIndex.Values, "frame inserted at wrong index");

// The same protection must work when the IList mutates at the wrong index
// and then throws an exception.
var misplacedAndThrowingValues = new FaultyList { 1.0, 2.0 };
misplacedAndThrowingValues.InsertAtFront = true;
misplacedAndThrowingValues.ThrowAfterInsert = true;
var wrongIndexThrow = new TestAnimation(new FaultyList { 10 }, misplacedAndThrowingValues);
Check(!KeyframeCollectionMutation.TryUpsert(wrongIndexThrow, 20, 5.0),
    "wrong-index insertion with exception was reported successful");
CheckUnchanged(wrongIndexThrow.Frames, wrongIndexThrow.Values,
    "wrong-index value insert threw after mutation");

// Immutable-style list properties can throw on re-read after initial discovery.
var flaky = new FlakyImmutableAnimation();
Check(!KeyframeCollectionMutation.TryUpsert(flaky, 20, 5.0),
    "throwing immutable property getter escaped the fallback");

// Integer-backed keyframes must not silently round 2.5 to 2 on replace
// or insert. Both failed operations must leave the original data intact.
var integerValues = new TestAnimation(new FaultyList { 10 }, new FaultyList { 1, 2 });
Check(!KeyframeCollectionMutation.TryUpsert(integerValues, 10, 2.5),
    "fractional integer-backed replacement was rounded");
CheckUnchanged(integerValues.Frames, integerValues.Values, "fractional replacement");
Check(!KeyframeCollectionMutation.TryUpsert(integerValues, 20, 2.5),
    "fractional integer-backed insertion was rounded");
CheckUnchanged(integerValues.Frames, integerValues.Values, "fractional insertion");

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
    public bool InsertAtFront { get; set; }
    public override void Insert(int index, object? value)
    {
        if (IgnoreInsert) return;
        base.Insert(InsertAtFront ? 0 : index, value);
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
