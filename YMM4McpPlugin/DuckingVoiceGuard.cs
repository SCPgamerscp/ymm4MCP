using System;
using System.Collections.Generic;
using System.Text.Json;

namespace YMM4McpPlugin
{
    internal sealed class DuckingVoiceGuard
    {
        private readonly bool _allVoices;
        private readonly Dictionary<string, (int frame, int length)> _expected;

        private DuckingVoiceGuard(bool allVoices,
            Dictionary<string, (int frame, int length)> expected)
        {
            _allVoices = allVoices;
            _expected = expected;
        }

        internal static DuckingVoiceGuard Parse(JsonElement selection)
        {
            if (selection.ValueKind != JsonValueKind.Object ||
                !selection.TryGetProperty("all_voices", out var allVoices) ||
                allVoices.ValueKind is not (JsonValueKind.True or JsonValueKind.False) ||
                !selection.TryGetProperty("items", out var items) ||
                items.ValueKind != JsonValueKind.Array || items.GetArrayLength() is < 1 or > 500)
                throw new ArgumentException("expected_ducking_voices must specify all_voices and 1..500 items");

            var expected = new Dictionary<string, (int frame, int length)>(StringComparer.Ordinal);
            foreach (var item in items.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object ||
                    !item.TryGetProperty("item_id", out var id) || id.ValueKind != JsonValueKind.String ||
                    !item.TryGetProperty("frame", out var start) || start.ValueKind != JsonValueKind.Number ||
                    !start.TryGetInt32(out int frame) || frame < 0 ||
                    !item.TryGetProperty("length", out var duration) || duration.ValueKind != JsonValueKind.Number ||
                    !duration.TryGetInt32(out int length) || length < 1 || (long)frame + length > int.MaxValue)
                    throw new ArgumentException("expected_ducking_voices contains an invalid item or timing");
                string itemId = KeyframeSelectorSafety.Parse(id.GetString(), null, null).itemId;
                if (!expected.TryAdd(itemId, (frame, length)))
                    throw new ArgumentException("expected_ducking_voices contains duplicate item IDs");
            }
            return new DuckingVoiceGuard(allVoices.GetBoolean(), expected);
        }

        internal bool Tracks(string itemId) => _allVoices || _expected.ContainsKey(itemId);

        internal bool Matches(IEnumerable<(string id, int? frame, int? length)> observed)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (id, frame, length) in observed)
            {
                if (!Tracks(id)) continue;
                if (!_expected.TryGetValue(id, out var expected) || !seen.Add(id) ||
                    frame != expected.frame || length != expected.length)
                    return false;
            }
            return seen.Count == _expected.Count;
        }

        // Match the single frame-zero key returned by ReadKeyframes, including
        // hosts that store the base volume outside the Values collection.
        internal static bool MatchesBaseVolume(double expected, IReadOnlyList<int> frames,
            IReadOnlyList<double> values, double? defaultValue)
            => (frames.Count == 0 && values.Count <= 1 ||
                frames.Count == 1 && frames[0] == 0 && values.Count == 1) &&
                (values.Count == 0 ? defaultValue : values[0]) == expected;
    }
}
