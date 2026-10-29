using System.ComponentModel;
using System.Globalization;
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
var initOnly = parameters.Single(p => p.GetProperty("name").GetString() == "InitOnly");
if (initOnly.GetProperty("writable").GetBoolean())
    throw new Exception("init-only property must not be advertised as editable");
var names = parameters.Select(p => p.GetProperty("name").GetString()!).ToArray();
if (!names.SequenceEqual(names.OrderBy(n => n, StringComparer.Ordinal)))
    throw new Exception("effect parameters must have stable ordinal ordering");
// Numeric metadata is serialized for AI consumption, not for local display.
// Changing the machine locale must not change its parseable decimal syntax.
var oldCulture = CultureInfo.CurrentCulture;
try
{
    CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
    using var french = JsonDocument.Parse(JsonSerializer.Serialize(EffectMetadata.Describe(typeof(FadeEffect))));
    var frenchParams = french.RootElement.GetProperty("parameters").EnumerateArray().ToArray();
    var gain = frenchParams.Single(p => p.GetProperty("name").GetString() == "Gain");
    if (gain.GetProperty("minimum").GetString() != "0.25"
        || gain.GetProperty("maximum").GetString() != "1.5"
        || gain.GetProperty("default_value").GetString() != "0.75")
        throw new Exception("numeric metadata must use invariant decimal separators");
    var invalid = frenchParams.Single(p => p.GetProperty("name").GetString() == "InvalidRange");
    if (invalid.GetProperty("minimum").ValueKind != JsonValueKind.Null
        || invalid.GetProperty("maximum").ValueKind != JsonValueKind.Null)
        throw new Exception("non-finite effect ranges must not be advertised as usable limits");
}
finally { CultureInfo.CurrentCulture = oldCulture; }
Console.WriteLine("Effect metadata tests passed");

sealed class FadeEffect
{
    [Range(0, 100), DefaultValue(50), DisplayName("強さ")]
    public int Amount { get; set; }
    public int Unknown { get; }
    public int InitOnly { get; init; }
    [Range(0.25, 1.5), DefaultValue(0.75)]
    public double Gain { get; set; }
    [Range(double.NegativeInfinity, double.PositiveInfinity)]
    public double InvalidRange { get; set; }
    public int Zeta { get; set; }
    public int alpha { get; set; }
}
