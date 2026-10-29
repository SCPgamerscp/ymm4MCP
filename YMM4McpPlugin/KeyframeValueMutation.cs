using System;
using System.Globalization;
using System.Reflection;

namespace YMM4McpPlugin
{
    // A reflected setter can silently clamp/ignore a value, or throw after
    // mutating it. Report success only after reading back the requested value.
    internal static class KeyframeValueMutation
    {
        internal static bool TrySet(object entry, double requested)
        {
            if (entry == null || !double.IsFinite(requested)) return false;
            var property = entry.GetType().GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            // Reflection can invoke private or init-only setters even though normal
            // YMM4 edits cannot. Never bypass the host's public mutability contract.
            if (property?.GetMethod?.IsPublic != true || property.SetMethod?.IsPublic != true
                || property.GetIndexParameters().Length != 0
                || property.SetMethod.ReturnParameter.GetRequiredCustomModifiers()
                    .Contains(typeof(System.Runtime.CompilerServices.IsExternalInit)))
                return false;
            var type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            var typeCode = Type.GetTypeCode(type);
            if (typeCode is not (TypeCode.Byte or TypeCode.SByte or TypeCode.Int16
                or TypeCode.UInt16 or TypeCode.Int32 or TypeCode.UInt32
                or TypeCode.Int64 or TypeCode.UInt64 or TypeCode.Single
                or TypeCode.Double or TypeCode.Decimal))
                return false;
            if (typeCode is TypeCode.Byte or TypeCode.SByte or TypeCode.Int16
                or TypeCode.UInt16 or TypeCode.Int32 or TypeCode.UInt32
                or TypeCode.Int64 or TypeCode.UInt64)
            {
                // Convert.ChangeType rounds fractional doubles for integer types.
                if (Math.Truncate(requested) != requested) return false;
            }

            object? previous;
            object converted;
            try
            {
                previous = property.GetValue(entry);
                converted = Convert.ChangeType(requested, type, CultureInfo.InvariantCulture);
                if (!double.IsFinite(Convert.ToDouble(converted, CultureInfo.InvariantCulture)))
                    return false; // e.g. a finite double overflowing a float.
            }
            catch { return false; }

            try
            {
                property.SetValue(entry, converted);
                var observed = property.GetValue(entry);
                if (observed != null)
                {
                    // Converting a float to text loses its exact widened value:
                    // float 0.1 becomes text "0.1", but widening it to double
                    // yields 0.10000000149011612. Compare numeric values directly.
                    var actual = Convert.ToDouble(observed, CultureInfo.InvariantCulture);
                    if (double.IsFinite(actual)
                        && actual == Convert.ToDouble(converted, CultureInfo.InvariantCulture))
                        return true;
                }
            }
            catch
            {
                // The host may have modified Value before throwing.
            }

            // Do not let a failed or clamped write corrupt an existing keyframe.
            // If restoration is unverifiable, stop the caller's fallback chain.
            try
            {
                property.SetValue(entry, previous);
                if (Equals(property.GetValue(entry), previous))
                    return false;
            }
            catch { }
            throw new InvalidOperationException(
                "YMM4 keyframe Value write failed and the original value could not be restored.");
        }
    }
}
