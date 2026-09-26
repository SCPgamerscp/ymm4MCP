using System.Text.Json;

namespace YMM4McpPlugin;

internal static class TimelineInputValidation
{
    internal static int GetInt(Dictionary<string, JsonElement> body, string key, int defaultValue)
    {
        if (!body.TryGetValue(key, out var value)) return defaultValue;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result)) return result;
        throw new ArgumentException($"{key} must be a 32-bit integer");
    }

    internal static int[]? GetLayers(Dictionary<string, JsonElement> body)
    {
        if (!body.TryGetValue("layers", out var value)) return null;
        if (value.ValueKind != JsonValueKind.Array) throw new ArgumentException("layers must be an array");
        var result = new List<int>();
        foreach (var element in value.EnumerateArray())
        {
            if (result.Count == 1000) throw new ArgumentException("layers must contain at most 1000 entries");
            if (element.ValueKind != JsonValueKind.Number || !element.TryGetInt32(out int layer) || layer < 0)
                throw new ArgumentException($"layers[{result.Count}] must be a non-negative 32-bit integer");
            result.Add(layer);
        }
        return result.ToArray();
    }

    internal static int AtLeast(int value, int minimum, string name)
    {
        if (value < minimum) throw new ArgumentException($"{name} must be at least {minimum}");
        return value;
    }

    internal static void RequireCoordinates(Dictionary<string, JsonElement> body)
    {
        if (!body.ContainsKey("frame") || !body.ContainsKey("layer"))
            throw new ArgumentException("frame and layer are required to identify an item");
        AtLeast(GetInt(body, "frame", -1), 0, "frame");
        AtLeast(GetInt(body, "layer", -1), 0, "layer");
    }

    internal static void RequireItemTarget(Dictionary<string, JsonElement> body)
    {
        if (body.TryGetValue("item_id", out var id))
        {
            if (id.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(id.GetString()))
                throw new ArgumentException("item_id must be a non-empty string");
            return;
        }
        RequireCoordinates(body);
    }

    internal static void RequireSelectionTarget(Dictionary<string, JsonElement> body)
    {
        if (body.TryGetValue("clear", out var clear) && clear.ValueKind == JsonValueKind.True) return;
        if (body.TryGetValue("item_id", out var id))
        {
            if (id.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(id.GetString()))
                throw new ArgumentException("item_id must be a non-empty string");
            return;
        }
        if (!body.ContainsKey("frame") && !body.ContainsKey("layer"))
            throw new ArgumentException("select requires item_id, frame, layer, or clear=true");
        if (body.ContainsKey("frame")) AtLeast(GetInt(body, "frame", -1), 0, "frame");
        if (body.ContainsKey("layer")) AtLeast(GetInt(body, "layer", -1), 0, "layer");
    }

    internal static int CheckPlacement(long frame, int length)
    {
        if (length <= 0 || frame < 0 || frame + length > int.MaxValue)
            throw new ArgumentException("item placement is outside the supported frame range");
        return (int)frame;
    }

    // Validate every target before invoking a setter. A bad later item must not move earlier items.
    internal static void ApplyShift<T>(IReadOnlyList<(T item, int frame, int length, int layer)> items,
        int delta, Action<T, int, int, int> apply)
    {
        var planned = items.Select(entry => (entry.item, entry.frame,
            target: CheckPlacement((long)entry.frame + delta, entry.length), entry.layer)).ToArray();
        foreach (var entry in planned) apply(entry.item, entry.frame, entry.target, entry.layer);
    }

    internal static void ApplyResolve<T>(IReadOnlyList<(T item, int frame, int length, int layer)> items,
        int gap, Action<T, int, int, int, int> apply)
    {
        AtLeast(gap, 0, nameof(gap));
        var planned = new List<(T item, int frame, int target, int length, int layer)>();
        foreach (var layerItems in items.GroupBy(entry => entry.layer))
        {
            long cursor = 0;
            foreach (var entry in layerItems.OrderBy(entry => entry.frame))
            {
                CheckPlacement(entry.frame, entry.length);
                int target = CheckPlacement(Math.Max((long)entry.frame, cursor), entry.length);
                planned.Add((entry.item, entry.frame, target, entry.length, entry.layer));
                cursor = (long)target + entry.length + gap;
            }
        }
        foreach (var entry in planned)
            if (entry.target != entry.frame) apply(entry.item, entry.frame, entry.target, entry.length, entry.layer);
    }
}
