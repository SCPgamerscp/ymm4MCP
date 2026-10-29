using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    /// <summary>Verify that an item created by a native YMM4 command is unambiguous
    /// before changing its properties or reporting a confirmed addition.</summary>
    internal static class AddedItemVerification
    {
        internal static bool IsExpectedType(object? item, string expectedName)
        {
            if (item == null || string.IsNullOrWhiteSpace(expectedName)) return false;
            for (Type? current = item.GetType(); current != null; current = current.BaseType)
            {
                if (string.Equals(current.Name, expectedName, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        internal static bool LengthMatches(object? actual, int expected) => actual switch
        {
            int n => n == expected,
            long n => n == expected,
            short n => n == expected,
            ushort n => n == expected,
            uint n => n == expected,
            ulong n => (ulong)expected == n,
            byte n => n == expected,
            sbyte n => n == expected,
            _ => false
        };

        internal static object? UniqueExpectedNewcomer(IEnumerable<object> newcomers, string expectedName)
        {
            object? candidate = null;
            foreach (var item in newcomers)
            {
                if (!IsExpectedType(item, expectedName)) continue;
                if (candidate != null) return null;
                candidate = item;
            }
            return candidate;
        }
    }
}
