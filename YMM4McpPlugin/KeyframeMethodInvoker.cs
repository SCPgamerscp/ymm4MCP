using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace YMM4McpPlugin
{
    // Keep the reflection fallback compatible with different YMM4 versions.
    // A boolean false is an explicit refusal, not a successful keyframe edit.
    internal static class KeyframeMethodInvoker
    {
        public static bool TryInvokeNamed(object target, IEnumerable<string> names, params object[] args)
        {
            if (target == null || names == null || args == null) return false;
            var methods = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (var name in names)
            {
                foreach (var method in methods.Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    // An asynchronous operation has not completed when reflection returns
                    // its Task. Reporting success here would silently lose a failed edit.
                    if (method.ContainsGenericParameters || IsAsyncReturn(method.ReturnType)) continue;
                    var parameters = method.GetParameters();
                    if (parameters.Length < args.Length || parameters.Any(p => p.ParameterType.IsByRef)) continue;
                    try
                    {
                        var call = new object?[parameters.Length];
                        bool compatible = true;
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            if (i < args.Length)
                            {
                                if (!TryConvertArgument(args[i], parameters[i].ParameterType, out call[i]))
                                {
                                    compatible = false;
                                    break;
                                }
                            }
                            else if (parameters[i].HasDefaultValue)
                                call[i] = parameters[i].DefaultValue;
                            else
                            {
                                compatible = false;
                                break;
                            }
                        }
                        if (!compatible) continue;
                        var result = method.Invoke(target, call);
                        if ((method.ReturnType == typeof(bool) || method.ReturnType == typeof(bool?))
                            && result is not true)
                            continue;
                        return true;
                    }
                    catch
                    {
                        // A different YMM4 version may expose an incompatible overload.
                    }
                }
            }
            return false;
        }

        private static bool IsAsyncReturn(Type type) =>
            typeof(Task).IsAssignableFrom(type)
            || type == typeof(ValueTask)
            || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(ValueTask<>));

        private static bool IsIntegerType(Type type) => Type.GetTypeCode(type) is
            TypeCode.Byte or TypeCode.SByte or TypeCode.Int16 or TypeCode.UInt16
            or TypeCode.Int32 or TypeCode.UInt32 or TypeCode.Int64 or TypeCode.UInt64;

        private static bool TryConvertArgument(object? argument, Type destination, out object? converted)
        {
            converted = null;
            var target = Nullable.GetUnderlyingType(destination) ?? destination;
            if (argument == null)
                return !destination.IsValueType || Nullable.GetUnderlyingType(destination) != null;
            // Even a directly assignable double/float must be finite. Otherwise
            // the fast path below would bypass the conversion validation.
            if (argument is double nonFiniteDouble && !double.IsFinite(nonFiniteDouble)
                || argument is float nonFiniteFloat && !float.IsFinite(nonFiniteFloat))
                return false;
            // Convert.ChangeType cannot convert an object to System.Object, a
            // derived class to its base class, or a value to Nullable<T>.
            if (target.IsInstanceOfType(argument))
            {
                converted = argument;
                return true;
            }
            if (target.IsEnum)
            {
                if (argument is string name && Enum.TryParse(target, name, false, out var parsed)
                    && parsed != null && Enum.IsDefined(target, parsed))
                {
                    converted = parsed;
                    return true;
                }
                return false;
            }
            if (argument is not IConvertible || !typeof(IConvertible).IsAssignableFrom(target))
                return false;
            try
            {
                // Convert.ChangeType rounds fractional doubles to integers. Never
                // silently move a keyframe to a different frame number.
                if (IsIntegerType(target))
                {
                    var numeric = Convert.ToDecimal(argument, CultureInfo.InvariantCulture);
                    if (numeric != decimal.Truncate(numeric)) return false;
                }
                converted = Convert.ChangeType(argument, target, CultureInfo.InvariantCulture);
                if (converted is double d && !double.IsFinite(d)
                    || converted is float f && !float.IsFinite(f))
                {
                    converted = null;
                    return false;
                }
                return true;
            }
            catch (Exception)
            {
                converted = null;
                return false;
            }
        }
    }
}
