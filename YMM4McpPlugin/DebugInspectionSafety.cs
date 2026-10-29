using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Input;

namespace YMM4McpPlugin
{
    /// <summary>Bounded, read-only helpers for diagnostic endpoints.</summary>
    internal static class DebugInspectionSafety
    {
        internal static Type[] RecoverLoadableTypes(ReflectionTypeLoadException error) =>
            error.Types.Where(type => type != null).Cast<Type>().ToArray();

        internal static Type[] LoadableTypes(Assembly assembly)
        {
            try { return assembly.GetTypes(); }
            catch (ReflectionTypeLoadException error) { return RecoverLoadableTypes(error); }
            catch { return Array.Empty<Type>(); }
        }

        internal static object[] SnapshotMenuItems(IEnumerable? items, int maxItems = 32)
        {
            if (items == null) return Array.Empty<object>();
            maxItems = Math.Clamp(maxItems, 1, 128);
            var result = new List<object>();
            foreach (var item in items)
            {
                if (item == null) continue;
                try
                {
                    var properties = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Select(p => new
                        {
                            p.Name,
                            TypeName = p.PropertyType.Name,
                            IsCommand = typeof(ICommand).IsAssignableFrom(p.PropertyType)
                        }).ToArray();
                    result.Add(new { type = item.GetType().Name, props = properties });
                }
                catch (Exception ex) when (ex is ReflectionTypeLoadException
                    or AmbiguousMatchException or TypeLoadException)
                {
                    result.Add(new { type = item.GetType().Name, error = "metadata unavailable" } as object);
                }
                if (result.Count >= maxItems) break;
            }
            return result.ToArray();
        }
    }
}
