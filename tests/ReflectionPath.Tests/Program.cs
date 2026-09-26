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
Console.WriteLine("Reflection path tests passed");

sealed class Item
{
    private readonly int privateValue = 42;
    public string Name { get; set; } = "example";
    public string? Optional { get; set; }
    public List<Item> Items { get; set; } = [];
}
