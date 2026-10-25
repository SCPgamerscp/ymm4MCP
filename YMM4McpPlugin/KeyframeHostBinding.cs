using System;
using System.Reflection;
using System.Linq;

namespace YMM4McpPlugin
{
    // Some YMM4 versions expose an item's shared KeyFrames collection before
    // individual Animation properties have been attached to it.  Mutating an
    // unbound Animation cannot create an item-level keyframe.
    internal static class KeyframeHostBinding
    {
        internal static bool HasConflictingBinding(object item, object animation)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance;
            try
            {
                var existing = animation.GetType().GetProperty("KeyFrames", flags)?.GetValue(animation);
                var shared = item.GetType().GetProperty("KeyFrames", flags)?.GetValue(item);
                return existing != null && shared != null && !ReferenceEquals(existing, shared);
            }
            catch
            {
                return false; // Unknown host layout: preserve existing fallback behavior.
            }
        }

        internal static bool TryBind(object item, object animation)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance;
            var animationProperty = animation.GetType().GetProperty("KeyFrames", flags);
            if (animationProperty == null) return false;
            object? existing;
            try { existing = animationProperty.GetValue(animation); }
            catch { return false; }
            // An Animation already attached to a different item's collection
            // must not be edited as though it belonged to this timeline item.
            if (existing != null) return !HasConflictingBinding(item, animation);

            object? shared;
            try { shared = item.GetType().GetProperty("KeyFrames", flags)?.GetValue(item); }
            catch { return false; }
            if (shared == null) return false;

            // Use the documented SetKeyFrames(KeyFrames) API rather than trying
            // to assign the getter-only Animation.KeyFrames property.
            // A YMM4 version may expose several compatible overloads, including
            // a broad object overload that rejects or ignores the call. Prefer
            // the exact collection type and keep trying if an overload fails.
            var methods = animation.GetType().GetMethods(flags)
                .Where(m => m.Name == "SetKeyFrames"
                    && m.GetParameters().Length == 1
                    && m.GetParameters()[0].ParameterType.IsInstanceOfType(shared))
                .OrderBy(m => m.GetParameters()[0].ParameterType == shared.GetType() ? 0 : 1);
            foreach (var method in methods)
            {
                try
                {
                    method.Invoke(animation, new[] { shared });
                    if (ReferenceEquals(animationProperty.GetValue(animation), shared))
                        return true;
                }
                catch
                {
                    // A rejected overload must not hide another usable one.
                }
            }
            return false;
        }
    }
}
