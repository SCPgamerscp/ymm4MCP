using System.Reflection;
using System.Globalization;

namespace YMM4McpPlugin;

internal static class EffectMetadata
{
    private static string? FormatAttributeValue(object? value)
    {
        // An AI client must see the same numeric limits regardless of the
        // editor's Windows locale (e.g. 0.5 rather than 0,5 in fr-FR).
        if (value is double d && !double.IsFinite(d)) return null;
        if (value is float f && !float.IsFinite(f)) return null;
        return value is IFormattable formattable
            ? formattable.ToString(null, CultureInfo.InvariantCulture)
            : value?.ToString();
    }

    internal static object Describe(Type effectType)
    {
        string? Value(PropertyInfo property, string attributeName, int index = 0)
        {
            var attribute = property.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name == attributeName);
            return attribute != null && attribute.ConstructorArguments.Count > index
                ? FormatAttributeValue(attribute.ConstructorArguments[index].Value) : null;
        }

        var parameters = effectType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.GetMethod?.IsPublic == true && p.GetIndexParameters().Length == 0)
            .OrderBy(p => p.Name, StringComparer.Ordinal)
            .Select(p =>
            {
                var range = p.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name == "RangeAttribute");
                int offset = range?.ConstructorArguments.Count == 3 ? 1 : 0;
                return new
                {
                    name = p.Name,
                    type = p.PropertyType.FullName ?? p.PropertyType.Name,
                    writable = p.SetMethod is { IsPublic: true } setter &&
                        !setter.ReturnParameter.GetRequiredCustomModifiers()
                            .Contains(typeof(System.Runtime.CompilerServices.IsExternalInit)),
                    minimum = range != null && range.ConstructorArguments.Count >= offset + 2
                        ? FormatAttributeValue(range.ConstructorArguments[offset].Value) : null,
                    maximum = range != null && range.ConstructorArguments.Count >= offset + 2
                        ? FormatAttributeValue(range.ConstructorArguments[offset + 1].Value) : null,
                    default_value = Value(p, "DefaultValueAttribute", 0),
                    display_name = Value(p, "DisplayNameAttribute"),
                    description = Value(p, "DescriptionAttribute"),
                    unit = Value(p, "UnitAttribute")
                };
            }).ToArray();
        return new { name = effectType.Name.EndsWith("Effect", StringComparison.Ordinal)
                    ? effectType.Name[..^6] : effectType.Name,
            full_name = effectType.FullName, parameters,
            note = "Ranges, defaults and units are returned only when declared as attributes; missing values are unknown." };
    }
}
