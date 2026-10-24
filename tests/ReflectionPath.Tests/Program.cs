using System.Collections;
using System.Collections.Immutable;
using YMM4McpPlugin;

static void Resolved(object source, string path, object? expected)
{
    if (!ReflectionPath.TryResolve(source, path, out var actual) || !Equals(actual, expected))
        throw new Exception($"Expected {path} to resolve to {expected}, got {actual}");
}

static void Missing(object source, string path)
{
    if (ReflectionPath.TryResolve(source, path, out _))
        throw new Exception($"Expected {path} to be missing");
}

var item = new Item { Items = [new Item()] };
Resolved(item, "Optional", null);
Resolved(item, "Items[0].Optional", null);
Resolved(item, "Items[0].Name", "example");
Resolved(item, "privateValue", 42);
Missing(item, "Unknown");
Missing(item, "Optional.Name");
Missing(item, "Items[1]");
Missing(item, "Items[bad]");
Missing(item, "Items[-1]");
Missing(item, "Items[0].Unknown");
Missing(item, "Items..Name");

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var implicitStart = new AnimationFixture(new List<int>(), new List<double> { 1.0 });
Check(KeyframeCollectionMutation.TryUpsert(implicitStart, 30, 2.0), "Insert first frame");
Check(KeyframeCollectionMutation.TryUpsert(implicitStart, 10, 1.5), "Insert sorted frame");
Check(KeyframeCollectionMutation.TryUpsert(implicitStart, 30, 3.0), "Update existing frame");
Check(KeyframeCollectionMutation.TryUpsert(implicitStart, 0, 0.5), "Update implicit start");
Check(((List<int>)implicitStart.KeyFrames.Frames).SequenceEqual(new[] { 10, 30 }), "Frame order");
Check(((List<double>)implicitStart.Values).SequenceEqual(new[] { 0.5, 1.5, 3.0 }), "Value alignment");

var explicitStart = new AnimationFixture(new List<int> { 0, 30 }, new List<double> { 1.0, 2.0 });
Check(KeyframeCollectionMutation.TryUpsert(explicitStart, 10, 1.5), "Insert explicit frame");
Check(((List<int>)explicitStart.KeyFrames.Frames).SequenceEqual(new[] { 0, 10, 30 }), "Explicit frame order");
Check(((List<double>)explicitStart.Values).SequenceEqual(new[] { 1.0, 1.5, 2.0 }), "Explicit values");

var firstWrapped = new WrappedValue { Value = 1.0 };
var wrapped = new AnimationFixture(new List<int>(), new List<WrappedValue> { firstWrapped });
Check(KeyframeCollectionMutation.TryUpsert(wrapped, 15, 2.0), "Insert wrapped value");
Check(KeyframeCollectionMutation.TryUpsert(wrapped, 0, 0.75), "Update wrapped start");
Check(ReferenceEquals(((List<WrappedValue>)wrapped.Values)[0], firstWrapped), "Preserve wrapped identity");
Check(((List<WrappedValue>)wrapped.Values)[1].Value == 2.0, "Wrapped inserted value");

var mismatched = new AnimationFixture(new List<int> { 10, 20 }, new List<double> { 1.0 });
Check(!KeyframeCollectionMutation.TryUpsert(mismatched, 15, 2.0), "Reject malformed collection");
Check(((List<int>)mismatched.KeyFrames.Frames).Count == 2, "Malformed frames unchanged");
Check(((List<double>)mismatched.Values).Count == 1, "Malformed values unchanged");

var fixedSize = new AnimationFixture(new[] { 10 }, new List<double> { 1.0, 2.0 });
Check(!KeyframeCollectionMutation.TryUpsert(fixedSize, 20, 3.0), "Reject fixed-size frames");
Check(((List<double>)fixedSize.Values).Count == 2, "Fixed-size failure is nonmutating");
Check(!KeyframeCollectionMutation.TryUpsert(implicitStart, -1, 1.0), "Reject negative frame");

// ImmutableList<T> implements non-generic IList but rejects in-place changes.
var immutable = new ImmutableAnimationFixture();
var oldFrames = immutable.KeyFrames.Frames;
var oldValues = immutable.Values;
Check(KeyframeCollectionMutation.TryUpsert(immutable, 20, 2.0), "Insert immutable frame");
Check(KeyframeCollectionMutation.TryUpsert(immutable, 10, 1.5), "Insert sorted immutable frame");
Check(KeyframeCollectionMutation.TryUpsert(immutable, 20, 3.0), "Update immutable frame");
Check(KeyframeCollectionMutation.TryUpsert(immutable, 0, 0.5), "Update immutable start");
var unchangedFrames = immutable.KeyFrames.Frames;
var unchangedValues = immutable.Values;
Check(KeyframeCollectionMutation.TryUpsert(immutable, 20, 3.0), "Idempotent immutable keyframe retry");
Check(KeyframeCollectionMutation.TryUpsert(immutable, 0, 0.5), "Idempotent immutable start retry");
Check(ReferenceEquals(immutable.KeyFrames.Frames, unchangedFrames)
    && ReferenceEquals(immutable.Values, unchangedValues),
    "Idempotent retry must not replace immutable collection references");
Check(immutable.KeyFrames.Frames.SequenceEqual(new[] { 10, 20 }), "Immutable frame order");
Check(immutable.Values.SequenceEqual(new[] { 0.5, 1.5, 3.0 }), "Immutable value alignment");
Check(oldFrames.Count == 0 && oldValues.SequenceEqual(new[] { 1.0 }),
    "Immutable originals must remain untouched");

var immutableReadOnly = new ImmutableReadOnlyAnimationFixture();
Check(!KeyframeCollectionMutation.TryUpsert(immutableReadOnly, 10, 2.0),
    "Reject immutable properties without setters");
Check(immutableReadOnly.KeyFrames.Frames.Count == 0
    && immutableReadOnly.Values.SequenceEqual(new[] { 1.0 }),
    "Read-only immutable failure leaves original unchanged");

var immutableMixed = new AnimationFixture(
    System.Collections.Immutable.ImmutableList<int>.Empty,
    new List<double> { 1.0 });
Check(!KeyframeCollectionMutation.TryUpsert(immutableMixed, 10, 2.0),
    "Reject mixed mutable and immutable layout");
Check(((List<double>)immutableMixed.Values).SequenceEqual(new[] { 1.0 }),
    "Mixed layout failure leaves values unchanged");

Console.WriteLine("Reflection path and keyframe collection tests passed");

sealed class Item
{
    private readonly int privateValue = 42;
    public string Name { get; set; } = "example";
    public string? Optional { get; set; }
    public List<Item> Items { get; set; } = [];
}

sealed class AnimationFixture(IList frames, IList values)
{
    public FrameFixture KeyFrames { get; } = new(frames);
    public IList Values { get; } = values;
}

sealed class FrameFixture(IList frames)
{
    public IList Frames { get; } = frames;
}

sealed class WrappedValue
{
    public double Value { get; set; }
}

sealed class ImmutableAnimationFixture
{
    public ImmutableFrameFixture KeyFrames { get; } = new();
    public ImmutableList<double> Values { get; set; } = ImmutableList.Create(1.0);
}

sealed class ImmutableFrameFixture
{
    public ImmutableList<int> Frames { get; set; } = ImmutableList<int>.Empty;
}

sealed class ImmutableReadOnlyAnimationFixture
{
    public ImmutableReadOnlyFrameFixture KeyFrames { get; } = new();
    public ImmutableList<double> Values { get; } = ImmutableList.Create(1.0);
}

sealed class ImmutableReadOnlyFrameFixture
{
    public ImmutableList<int> Frames { get; } = ImmutableList<int>.Empty;
}
