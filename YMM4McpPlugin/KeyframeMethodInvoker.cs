using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace YMM4McpPlugin
{
    // Keep the reflection fallback compatible with different YMM4 versions.
    // A boolean false is an explicit refusal, not a successful keyframe edit.
    internal static class KeyframeMethodInvoker
    {
        public static bool TryInvokeNamed(object target, IEnumerable<string> names, params object[] args)
        {
            var methods = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (var name in names)
            {
                foreach (var method in methods.Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    var parameters = method.GetParameters();
                    if (parameters.Length < args.Length) continue;
                    try
                    {
                        var call = new object?[parameters.Length];
                        bool compatible = true;
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            if (i < args.Length)
                                call[i] = Convert.ChangeType(args[i], parameters[i].ParameterType, CultureInfo.InvariantCulture);
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
    }
}
