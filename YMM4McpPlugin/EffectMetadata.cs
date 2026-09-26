using System.Reflection;

namespace YMM4McpPlugin;

internal static class EffectMetadata
{
    internal static object Describe(Type effectType)
    {
        string? Value(PropertyInfo property, string attributeName, int index = 0)
        {
            var attribute = property.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name == attributeName);
            return attribute != null && attribute.ConstructorArguments.Count > index
                ? attribute.ConstructorArguments[index].Value?.ToString() : null;
        }

        var parameters = effectType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.GetMethod?.IsPublic == true && p.GetIndexParameters().Length == 0)
            .OrderBy(p => p.Name)
            .Select(p =>
            {
                var range = p.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name == "RangeAttribute");
                int offset = range?.ConstructorArguments.Count == 3 ? 1 : 0;
                return new
                {
                    name = p.Name,
                    type = p.PropertyType.FullName ?? p.PropertyType.Name,
                    writable = p.SetMethod?.IsPublic == true,
                    minimum = range != null && range.ConstructorArguments.Count >= offset + 2
                        ? range.ConstructorArguments[offset].Value?.ToString() : null,
                    maximum = range != null && range.ConstructorArguments.Count >= offset + 2
                        ? range.ConstructorArguments[offset + 1].Value?.ToString() : null,
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
