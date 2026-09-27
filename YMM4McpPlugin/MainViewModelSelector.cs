using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    internal static class MainViewModelSelector
    {
        // YMM4's Index is a window/order detail and has changed between releases.
        // Accept only a real MainViewModel, then prefer the active editing window.
        internal static object? Select(IEnumerable<(object? Context, bool Active, bool Visible)> windows)
        {
            object? visible = null, other = null;
            foreach (var (context, active, isVisible) in windows)
            {
                if (context == null || !IsMainViewModel(context.GetType())) continue;
                if (active) return context;
                if (isVisible) visible ??= context;
                else other ??= context;
            }
            return visible ?? other;
        }

        private static bool IsMainViewModel(Type type)
        {
            for (Type? current = type; current != null; current = current.BaseType)
                if (current.Name == "MainViewModel") return true;
            return false;
        }
    }
}
