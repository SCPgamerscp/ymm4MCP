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
                // A throwing host getter makes ownership unknowable. Never
                // mutate an animation that may belong to another item.
                return true;
            }
        }

        // Never report a nonzero keyframe insertion as successful when the
        // animation could not be attached to the item's shared KeyFrames.
        // Frame zero does not create a new shared keyframe.
        internal static bool CanSetAt(object item, object animation, int relativeFrame)
            => relativeFrame == 0 || TryBind(item, animation);

        internal static bool TryBind(object item, object animation)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance;
            var animationProperty = animation.GetType().GetProperty("KeyFrames", flags);
            if (animationProperty == null) return false;
            object? existing;
            try { existing = animationProperty.GetValue(animation); }
            catch { return false; }
            // A non-null animation binding is only valid if it is the exact
            // shared collection owned by the requested timeline item.
            object? shared;
            try { shared = item.GetType().GetProperty("KeyFrames", flags)?.GetValue(item); }
            catch { return false; }
            if (shared == null) return false;
            if (existing != null) return ReferenceEquals(existing, shared);

            // Use the documented SetKeyFrames(KeyFrames) API rather than trying
            // to assign the getter-only Animation.KeyFrames property.
            // A YMM4 version may expose several compatible overloads, including
            // a broad object overload that rejects or ignores the call. Prefer
            // the exact collection type and keep trying if an overload fails.
            var methods = animation.GetType().GetMethods(flags)
                .Where(m => m.Name == "SetKeyFrames"
                    && m.GetParameters().Length == 1
                    && m.GetParameters()[0].ParameterType.IsInstanceOfType(shared)
                    // This synchronous binding path cannot await Task/ValueTask.
                    // Invoking such a method could report success before binding completes.
                    && !typeof(System.Threading.Tasks.Task).IsAssignableFrom(m.ReturnType)
                    && m.ReturnType != typeof(System.Threading.Tasks.ValueTask)
                    && !(m.ReturnType.IsGenericType && m.ReturnType.GetGenericTypeDefinition()
                        == typeof(System.Threading.Tasks.ValueTask<>)))
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
