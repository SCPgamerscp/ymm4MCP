using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using YMM4McpPlugin;

using var document = JsonDocument.Parse(JsonSerializer.Serialize(EffectMetadata.Describe(typeof(FadeEffect))));
var root = document.RootElement;
if (root.GetProperty("name").GetString() != "Fade") throw new Exception("effect name");
var parameters = root.GetProperty("parameters").EnumerateArray().ToArray();
var amount = parameters.Single(p => p.GetProperty("name").GetString() == "Amount");
if (amount.GetProperty("minimum").GetString() != "0" || amount.GetProperty("maximum").GetString() != "100"
    || amount.GetProperty("default_value").GetString() != "50"
    || amount.GetProperty("display_name").GetString() != "強さ"
    || amount.GetProperty("writable").GetBoolean() != true) throw new Exception("effect attributes");
var unknown = parameters.Single(p => p.GetProperty("name").GetString() == "Unknown");
if (unknown.GetProperty("minimum").ValueKind != JsonValueKind.Null || unknown.GetProperty("writable").GetBoolean())
    throw new Exception("unknown metadata must remain unknown");
Console.WriteLine("Effect metadata tests passed");

sealed class FadeEffect
{
    [Range(0, 100), DefaultValue(50), DisplayName("強さ")]
    public int Amount { get; set; }
    public int Unknown { get; }
}
