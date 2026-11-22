using System;

namespace YMM4McpPlugin
{
    internal static class FacePathMutation
    {
        internal static string SetAndVerify(string expected, Action set, Func<object?> read)
        {
            set();
            if (read() is string observed && string.Equals(observed, expected, StringComparison.Ordinal))
                return observed;
            throw new InvalidOperationException("FacePath readback did not match the requested value");
        }
    }
}
