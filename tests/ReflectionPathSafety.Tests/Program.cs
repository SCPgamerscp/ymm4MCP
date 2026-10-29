using System.Collections;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

Check(ReflectionPath.TryResolve(new Root { Data = new IndexedOnlyList() }, "Data[0]", out var value)
    && Equals(value, "first"), "IList must resolve by index without enumerating");
Check(!ReflectionPath.TryResolve(new Root { Data = new IndexedOnlyList() }, "Data[1]", out _),
    "Out-of-range IList indices must be rejected");
Check(!ReflectionPath.TryResolve(new Root { Data = new IndexedOnlyList(throwIndexer: true) }, "Data[0]", out _),
    "Throwing IList indexers must not escape");
Check(!ReflectionPath.TryResolve(new Root { Data = new ThrowingEnumerable() }, "Data[0]", out _),
    "Throwing IEnumerable must not escape");
Check(!ReflectionPath.TryResolve(new Root { Data = new EndlessEnumerable() }, "Data[100001]", out _),
    "Unbounded lazy enumeration must be rejected before iteration");
Check(ReflectionPath.TryResolve(new Root { Data = new[] { 7, 8, 9 } }, "Data[2]", out value)
    && Equals(value, 9), "Normal array indexing must continue to work");
Console.WriteLine("Reflection path safety regression tests passed");

sealed class Root
{
    public object Data { get; init; } = null!;
}

sealed class ThrowingEnumerable : IEnumerable
{
    public IEnumerator GetEnumerator() => throw new InvalidOperationException("host collection is unavailable");
}

sealed class EndlessEnumerable : IEnumerable
{
    public IEnumerator GetEnumerator()
    {
        while (true) yield return 1;
    }
}

sealed class IndexedOnlyList(bool throwIndexer = false) : IList
{
    public int Count => 1;
    public bool IsSynchronized => false;
    public object SyncRoot => this;
    public bool IsFixedSize => true;
    public bool IsReadOnly => true;
    public object? this[int index]
    {
        get
        {
            if (throwIndexer) throw new InvalidOperationException("host indexer failed");
            return index == 0 ? "first" : throw new ArgumentOutOfRangeException(nameof(index));
        }
        set => throw new NotSupportedException();
    }
    public IEnumerator GetEnumerator() => throw new InvalidOperationException("Enumeration must not be used");
    public void CopyTo(Array array, int index) => throw new NotSupportedException();
    public int Add(object? value) => throw new NotSupportedException();
    public void Clear() => throw new NotSupportedException();
    public bool Contains(object? value) => false;
    public int IndexOf(object? value) => -1;
    public void Insert(int index, object? value) => throw new NotSupportedException();
    public void Remove(object? value) => throw new NotSupportedException();
    public void RemoveAt(int index) => throw new NotSupportedException();
}
