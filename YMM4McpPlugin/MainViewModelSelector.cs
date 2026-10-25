using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    internal static class MainViewModelSelector
    {
        // YMM4's Index is a window/order detail and has changed between releases.
        // Accept only a real MainViewModel, then prefer the active editing window.
        // Never choose an arbitrary editor when several windows share the same
        // priority: doing so could apply an AI edit to the wrong project.
        internal static object? Select(IEnumerable<(object? Context, bool Active, bool Visible)> windows)
        {
            object? active = null, visible = null, other = null;
            bool multipleActive = false, multipleVisible = false, multipleOther = false;
            if (windows == null) return null;
            try
            {
                foreach (var (context, isActive, isVisible) in windows)
                {
                    if (context == null || !IsMainViewModel(context.GetType())) continue;
                    if (isActive)
                        Consider(context, ref active, ref multipleActive);
                    else if (isVisible)
                        Consider(context, ref visible, ref multipleVisible);
                    else
                        Consider(context, ref other, ref multipleOther);
                }
            }
            catch (InvalidOperationException)
            {
                // Window collections can change while a YMM4 editor is closing.
                // A partial enumeration is not safe evidence of a unique editor.
                return null;
            }
            if (active != null) return multipleActive ? null : active;
            if (visible != null) return multipleVisible ? null : visible;
            return multipleOther ? null : other;
        }

        private static void Consider(object context, ref object? candidate, ref bool ambiguous)
        {
            if (candidate == null) candidate = context;
            else if (!ReferenceEquals(candidate, context)) ambiguous = true;
        }

        private static bool IsMainViewModel(Type type)
        {
            for (Type? current = type; current != null; current = current.BaseType)
                if (current.Name == "MainViewModel") return true;
            return false;
        }
    }
}
