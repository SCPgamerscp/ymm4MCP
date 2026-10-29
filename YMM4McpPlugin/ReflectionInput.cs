using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace YMM4McpPlugin
{
    /// <summary>Bind untrusted JSON to reflection arguments without lossy numeric conversions
    /// or silent string fallbacks. No method is invoked by this helper.</summary>
    internal static class ReflectionInput
    {
        internal static bool TryConvert(JsonElement input, Type requestedType, out object? value)
        {
            value = null;
            var type = Nullable.GetUnderlyingType(requestedType) ?? requestedType;
            if (input.ValueKind == JsonValueKind.Null)
                return !requestedType.IsValueType || Nullable.GetUnderlyingType(requestedType) != null;
            if (input.ValueKind == JsonValueKind.Undefined) return false;

            if (type == typeof(JsonElement) || type == typeof(object))
            {
                value = input.Clone();
                return true;
            }
            if (type == typeof(string))
            {
                if (input.ValueKind != JsonValueKind.String) return false;
                value = input.GetString();
                return true;
            }
            if (type == typeof(bool))
            {
                if (input.ValueKind == JsonValueKind.True || input.ValueKind == JsonValueKind.False)
                {
                    value = input.GetBoolean();
                    return true;
                }
                if (input.ValueKind == JsonValueKind.String
                    && bool.TryParse(input.GetString(), out var parsedBool))
                {
                    value = parsedBool;
                    return true;
                }
                return false;
            }
            if (type.IsEnum)
            {
                if (input.ValueKind == JsonValueKind.String)
                {
                    try
                    {
                        value = Enum.Parse(type, input.GetString()!, ignoreCase: true);
                        return true;
                    }
                    catch (ArgumentException) { return false; }
                    catch (OverflowException) { return false; }
                }
                if (TryConvert(input, Enum.GetUnderlyingType(type), out var enumNumber)
                    && enumNumber != null)
                {
                    value = Enum.ToObject(type, enumNumber);
                    return true;
                }
                return false;
            }
            if (type == typeof(Guid) && input.ValueKind == JsonValueKind.String
                && Guid.TryParse(input.GetString(), out var guid))
            {
                value = guid;
                return true;
            }
            if (type == typeof(DateTime) && input.ValueKind == JsonValueKind.String
                && DateTime.TryParse(input.GetString(), CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind, out var date))
            {
                value = date;
                return true;
            }
            if (type == typeof(TimeSpan) && input.ValueKind == JsonValueKind.String
                && TimeSpan.TryParse(input.GetString(), CultureInfo.InvariantCulture, out var span))
            {
                value = span;
                return true;
            }
            if (type == typeof(char) && input.ValueKind == JsonValueKind.String)
            {
                var s = input.GetString();
                if (s?.Length == 1) { value = s[0]; return true; }
                return false;
            }
            if (input.ValueKind != JsonValueKind.Number && input.ValueKind != JsonValueKind.String)
                return false;

            var raw = input.ValueKind == JsonValueKind.Number ? input.GetRawText() : input.GetString();
            if (string.IsNullOrWhiteSpace(raw)) return false;
            if (type == typeof(double) || type == typeof(float))
            {
                if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)
                    || !double.IsFinite(d)) return false;
                if (type == typeof(double)) { value = d; return true; }
                var f = (float)d;
                if (!float.IsFinite(f)) return false;
                value = f;
                return true;
            }
            if (!decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture,
                out var number)) return false;
            if (type == typeof(decimal)) { value = number; return true; }
            if (decimal.Truncate(number) != number) return false;
            if (type != typeof(byte) && type != typeof(sbyte)
                && type != typeof(short) && type != typeof(ushort)
                && type != typeof(int) && type != typeof(uint)
                && type != typeof(long) && type != typeof(ulong)) return false;
            try
            {
                value = Convert.ChangeType(number, type, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex) when (ex is OverflowException or InvalidCastException or FormatException)
            {
                return false;
            }
        }

        internal static bool TryBind(MethodInfo method, IReadOnlyList<JsonElement> input, out object?[] arguments)
        {
            arguments = Array.Empty<object?>();
            if (method.ContainsGenericParameters) return false;
            var parameters = method.GetParameters();
            if (parameters.Length != input.Count
                || parameters.Any(p => p.IsOut || p.ParameterType.IsByRef || p.ParameterType.IsPointer))
                return false;
            var converted = new object?[parameters.Length];
            for (var i = 0; i < parameters.Length; i++)
            {
                if (!TryConvert(input[i], parameters[i].ParameterType, out converted[i]))
                    return false;
            }
            arguments = converted;
            return true;
        }

        /// <summary>Never guess between compatible overloads: the wrong choice may mutate a project.</summary>
        internal static bool TrySelectMethod(IEnumerable<MethodInfo> methods,
            IReadOnlyList<JsonElement> input, out MethodInfo? selected,
            out object?[]? arguments, out bool ambiguous)
        {
            selected = null;
            arguments = null;
            ambiguous = false;
            foreach (var method in methods)
            {
                if (!TryBind(method, input, out var bound)) continue;
                if (selected != null)
                {
                    selected = null;
                    arguments = null;
                    ambiguous = true;
                    return false;
                }
                selected = method;
                arguments = bound;
            }
            return selected != null;
        }
    }
}
