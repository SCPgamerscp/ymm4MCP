using System;
using System.Collections;
using System.Collections.Immutable;
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

// Some IList implementations corrupt an existing value without growing the
// collection, or insert more than one value. These states cannot be restored
// by undoing one insertion: fail loudly instead of allowing another fallback.
var overwrittenValues = new FaultyList { 1.0, 2.0 };
overwrittenValues.ReplaceOnInsert = true;
var corruptValues = new TestAnimation(new FaultyList { 10 }, overwrittenValues);
bool rejectedCorruptValues = false;
try { KeyframeCollectionMutation.TryUpsert(corruptValues, 20, 5.0); }
catch (InvalidOperationException) { rejectedCorruptValues = true; }
Check(rejectedCorruptValues,
    "same-count value corruption must not be reported as an ordinary fallback");

var doubleValues = new FaultyList { 1.0, 2.0 };
doubleValues.InsertTwice = true;
var doubleInsert = new TestAnimation(new FaultyList { 10 }, doubleValues);
bool rejectedDoubleInsert = false;
try { KeyframeCollectionMutation.TryUpsert(doubleInsert, 20, 5.0); }
catch (InvalidOperationException) { rejectedDoubleInsert = true; }
Check(rejectedDoubleInsert,
    "two inserted values cannot be silently accepted or rolled back as one");

var overwrittenFrames = new FaultyList { 10 };
overwrittenFrames.ReplaceOnInsert = true;
var corruptFrames = new TestAnimation(overwrittenFrames, new FaultyList { 1.0, 2.0 });
bool rejectedCorruptFrames = false;
try { KeyframeCollectionMutation.TryUpsert(corruptFrames, 20, 5.0); }
catch (InvalidOperationException) { rejectedCorruptFrames = true; }
Check(rejectedCorruptFrames,
    "same-count frame corruption must not be reported as a safe fallback");
Check(corruptFrames.Values.Count == 2 && Convert.ToDouble(corruptFrames.Values[0]) == 1.0
    && Convert.ToDouble(corruptFrames.Values[1]) == 2.0,
    "unrecoverable frame insertion must still restore the values collection");

// A collection can throw *after* RemoveAt has successfully undone Insert.
// A verified original snapshot is still a safe failure.
var removeThrows = new FaultyList { 1.0, 2.0 };
removeThrows.ThrowAfterInsert = true;
removeThrows.ThrowAfterRemove = true;
var recoverableRemoveThrow = new TestAnimation(new FaultyList { 10 }, removeThrows);
Check(!KeyframeCollectionMutation.TryUpsert(recoverableRemoveThrow, 20, 5.0),
    "verified rollback after RemoveAt throws must remain a safe failure");
CheckUnchanged(recoverableRemoveThrow.Frames, recoverableRemoveThrow.Values,
    "remove throws after restoring original values");

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

// A reflected Value setter may mutate and then throw. The failed replacement
// must restore the old keyframe value before reporting failure.
var throwsOnce = new ThrowingKeyframeValue(2.0) { ThrowOnceAfterWrite = true };
var recoverable = new TestAnimation(
    new FaultyList { 10 }, new FaultyList { new ThrowingKeyframeValue(1.0), throwsOnce });
Check(!KeyframeCollectionMutation.TryUpsert(recoverable, 10, 9.0),
    "setter throwing after mutation was reported successful");
Check(throwsOnce.Value == 2.0,
    "setter throwing after mutation did not restore the original keyframe value");

// If the original value cannot be verified after attempted restoration,
// abort the edit rather than continuing into a second fallback.
var alwaysThrows = new ThrowingKeyframeValue(2.0) { ThrowEveryWrite = true };
var unrecoverable = new TestAnimation(
    new FaultyList { 10 }, new FaultyList { new ThrowingKeyframeValue(1.0), alwaysThrows });
bool refusedUnsafeFallback = false;
try { KeyframeCollectionMutation.TryUpsert(unrecoverable, 10, 9.0); }
catch (InvalidOperationException) { refusedUnsafeFallback = true; }
Check(refusedUnsafeFallback,
    "unverifiable reflected Value restoration must stop the fallback chain");

// Immutable collections use functional Insert/SetItem followed by two host
// property assignments. Both assignments must commit or both must roll back.
var immutable = new TestAnimation(
    ImmutableList.Create(10), ImmutableList.Create(1.0, 2.0));
Check(KeyframeCollectionMutation.TryUpsert(immutable, 20, 5.0),
    "valid immutable keyframe insertion should succeed");
Check(immutable.Frames.Count == 2 && Convert.ToInt32(immutable.Frames[1]) == 20
    && immutable.Values.Count == 3 && Convert.ToDouble(immutable.Values[2]) == 5.0,
    "immutable insertion did not commit both collections");

// A host Frames setter can install the new immutable list and then throw,
// then reject the rollback. Returning false would let a later fallback edit
// the partially mutated animation, so this must throw instead.
var unsafeImmutable = new FailingImmutableAnimation(
    ImmutableList.Create(10), ImmutableList.Create(1.0, 2.0));
bool rejectedUnrestoredImmutable = false;
try { KeyframeCollectionMutation.TryUpsert(unsafeImmutable, 20, 5.0); }
catch (InvalidOperationException) { rejectedUnrestoredImmutable = true; }
Check(rejectedUnrestoredImmutable,
    "unrestorable immutable frame assignment must stop fallback edits");
Check(unsafeImmutable.Frames.Count == 2 && unsafeImmutable.Values.Count == 2,
    "immutable rollback test did not exercise a partially committed frame update");

// Float-backed values must be verified through the same conversion as the
// actual collection entry: Convert.ToDouble(0.1f) is not equal to 0.1.
var floatMutable = new TestAnimation(
    new FaultyList { 10 }, new FaultyList { 1.0f, 2.0f });
Check(KeyframeCollectionMutation.TryUpsert(floatMutable, 20, 0.1),
    "float-backed mutable insertion should succeed");
Check(floatMutable.Values.Count == 3
    && floatMutable.Values[2] is float insertedFloat && insertedFloat == 0.1f,
    "float-backed mutable insertion wrote the wrong value");
Check(KeyframeCollectionMutation.TryUpsert(floatMutable, 10, 0.1),
    "float-backed mutable replacement should succeed");
Check(floatMutable.Values[1] is float replacedFloat && replacedFloat == 0.1f,
    "float-backed mutable replacement wrote the wrong value");

var floatProperty = new FloatKeyframeValue { Value = 2.0f };
var floatPropertyAnimation = new TestAnimation(
    new FaultyList { 10 },
    new FaultyList { new FloatKeyframeValue { Value = 1.0f }, floatProperty });
Check(KeyframeCollectionMutation.TryUpsert(floatPropertyAnimation, 10, 0.1),
    "float-backed Value property replacement should succeed");
Check(floatProperty.Value == 0.1f,
    "float-backed Value property replacement wrote the wrong value");

var floatImmutable = new TestAnimation(
    ImmutableList.Create(10), ImmutableList.Create(1.0f, 2.0f));
Check(KeyframeCollectionMutation.TryUpsert(floatImmutable, 20, 0.1),
    "float-backed immutable insertion should succeed");
Check(floatImmutable.Values.Count == 3
    && floatImmutable.Values[2] is float immutableInserted
    && immutableInserted == 0.1f,
    "float-backed immutable insertion wrote the wrong value");
Check(KeyframeCollectionMutation.TryUpsert(floatImmutable, 10, 0.1),
    "float-backed immutable replacement should succeed");
Check(floatImmutable.Values[1] is float immutableReplaced
    && immutableReplaced == 0.1f,
    "float-backed immutable replacement wrote the wrong value");

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
    public bool ReplaceOnInsert { get; set; }
    public bool InsertTwice { get; set; }
    public bool ThrowAfterRemove { get; set; }
    public override void Insert(int index, object? value)
    {
        if (IgnoreInsert) return;
        if (ReplaceOnInsert) this[index == Count ? Count - 1 : index] = value;
        else
        {
            base.Insert(InsertAtFront ? 0 : index, value);
            if (InsertTwice) base.Insert(InsertAtFront ? 0 : index, value);
        }
        if (ThrowAfterInsert) throw new InvalidOperationException("insert failed after mutation");
    }
    public override void RemoveAt(int index)
    {
        base.RemoveAt(index);
        if (ThrowAfterRemove) throw new InvalidOperationException("remove failed after mutation");
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

internal sealed class FailingImmutableAnimation
{
    private IList frames;
    private IList values;
    private int frameWrites;
    public FailingImmutableAnimation(IList frames, IList values)
    {
        this.frames = frames;
        this.values = values;
    }
    public IList Frames
    {
        get => frames;
        set
        {
            frameWrites++;
            if (frameWrites == 1)
            {
                frames = value;
                throw new InvalidOperationException("frame setter threw after commit");
            }
            throw new InvalidOperationException("frame setter refuses rollback");
        }
    }
    public IList Values { get => values; set => values = value; }
}

internal sealed class ThrowingKeyframeValue
{
    private double value;
    public ThrowingKeyframeValue(double initial) { value = initial; }
    public bool ThrowOnceAfterWrite { get; set; }
    public bool ThrowEveryWrite { get; set; }
    public double Value
    {
        get => value;
        set
        {
            this.value = value;
            if (ThrowOnceAfterWrite)
            {
                ThrowOnceAfterWrite = false;
                throw new InvalidOperationException("setter threw after mutation");
            }
            if (ThrowEveryWrite)
                throw new InvalidOperationException("setter always throws after mutation");
        }
    }
}

internal sealed class FloatKeyframeValue
{
    public float Value { get; set; }
}
