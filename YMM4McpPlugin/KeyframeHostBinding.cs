using System;
using System.Reflection;

namespace YMM4McpPlugin
{
    // Some YMM4 versions expose an item's shared KeyFrames collection before
    // individual Animation properties have been attached to it.  Mutating an
    // unbound Animation cannot create an item-level keyframe.
    internal static class KeyframeHostBinding
    {
        internal static bool TryBind(object item, object animation)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance;
            var animationProperty = animation.GetType().GetProperty("KeyFrames", flags);
            if (animationProperty == null) return false;
            object? existing;
            try { existing = animationProperty.GetValue(animation); }
            catch { return false; }
            if (existing != null) return true;

            object? shared;
            try { shared = item.GetType().GetProperty("KeyFrames", flags)?.GetValue(item); }
            catch { return false; }
            if (shared == null) return false;

            // Use the documented SetKeyFrames(KeyFrames) API rather than trying
            // to assign the getter-only Animation.KeyFrames property.
            foreach (var method in animation.GetType().GetMethods(flags))
            {
                if (method.Name != "SetKeyFrames") continue;
                var parameters = method.GetParameters();
                if (parameters.Length != 1 || !parameters[0].ParameterType.IsInstanceOfType(shared))
                    continue;
                try
                {
                    method.Invoke(animation, new[] { shared });
                    return ReferenceEquals(animationProperty.GetValue(animation), shared);
                }
                catch { return false; }
            }
            return false;
        }
    }
}
