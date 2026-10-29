using System.Collections;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
Check(KeyframeInsertionSafety.IsComplete(new[] { 10, 20 }, 3,
    new[] { 10, 15, 20 }, 4, 15), "valid implicit-start insertion");
Check(KeyframeInsertionSafety.IsComplete(new[] { 10, 20 }, 2,
    new[] { 10, 15, 20 }, 3, 15), "valid no-start insertion");
Check(!KeyframeInsertionSafety.IsComplete(new[] { 20, 10 }, 2,
    new[] { 15, 20, 10 }, 3, 15), "unordered source must not pass");
Check(!KeyframeInsertionSafety.IsComplete(new[] { 10, 10 }, 2,
    new[] { 5, 10, 10 }, 3, 5), "duplicate source frames must not pass");
Check(!KeyframeInsertionSafety.IsComplete(new[] { 10, 20 }, 2,
    new[] { 10, 15, 15 }, 3, 15), "duplicate resulting frames must not pass");
Check(!KeyframeInsertionSafety.IsComplete(new[] { -1, 10 }, 2,
    new[] { -1, 5, 10 }, 3, 5), "negative source frames must not pass");
Check(!KeyframeInsertionSafety.IsComplete(new[] { 10, 20 }, 2,
    new[] { 10, 20, 15 }, 3, 15), "reordered resulting frames must not pass");
Check(!KeyframeInsertionSafety.IsComplete(new BrokenFrames(), 1,
    new[] { 5, 10 }, 2, 5), "throwing host frame collection must fail closed");
var values = new ArrayList { 1, 2, 3 };
Check(KeyframeInsertionSafety.TryRemoveInsertedValue(values, 2, 1) && values.Count == 2,
    "valid rollback removes the inserted value");
Check(!KeyframeInsertionSafety.TryRemoveInsertedValue(new ThrowingList(), 1, 0),
    "throwing host IList metadata must fail closed");
Check(!KeyframeInsertionSafety.TryRemoveInsertedValue(ArrayList.ReadOnly(new ArrayList { 1, 2 }), 1, 1),
    "read-only value collections must not be mutated");
Console.WriteLine("Keyframe insertion snapshot integrity tests passed");

sealed class BrokenFrames : IReadOnlyList<int>
{
    public int Count => 1;
    public int this[int index] => throw new InvalidOperationException("host getter failed");
    public IEnumerator<int> GetEnumerator() => throw new InvalidOperationException("host enumeration failed");
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

sealed class ThrowingList : IList
{
    public int Count => throw new InvalidOperationException("host Count failed");
    public bool IsReadOnly => false;
    public bool IsFixedSize => false;
    public bool IsSynchronized => false;
    public object SyncRoot => this;
    public object? this[int index] { get => null; set { } }
    public int Add(object? value) => 0;
    public void Clear() { }
    public bool Contains(object? value) => false;
    public int IndexOf(object? value) => -1;
    public void Insert(int index, object? value) { }
    public void Remove(object? value) { }
    public void RemoveAt(int index) { }
    public void CopyTo(Array array, int index) { }
    public IEnumerator GetEnumerator() => Array.Empty<object>().GetEnumerator();
}
