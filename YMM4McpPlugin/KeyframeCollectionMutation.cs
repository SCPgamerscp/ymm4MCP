using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace YMM4McpPlugin
{
    // For YMM4 versions with mutable Frames/Values but no AddKeyFrame API.
    // Unknown collection layouts are rejected rather than modified speculatively.
    internal static class KeyframeCollectionMutation
    {
        internal static bool TryUpsert(object animation, int frame, double value)
        {
            if (frame < 0 || !double.IsFinite(value)) return false;
            var owner = Read(animation, "KeyFrames") ?? animation;
            if ((Read(owner, "Frames") ?? Read(owner, "FrameList")) is not IList frames
                || Read(animation, "Values") is not IList values)
                return false;
            if (frames.IsReadOnly || frames.IsFixedSize || values.IsReadOnly || values.IsFixedSize)
                return false;

            int offset;
            if (values.Count == frames.Count + 1) offset = 1;
            else if (values.Count == frames.Count && frames.Count > 0) offset = 0;
            else return false;

            int insertAt = frames.Count;
            int previous = -1;
            for (int i = 0; i < frames.Count; i++)
            {
                if (!int.TryParse(Convert.ToString(frames[i], CultureInfo.InvariantCulture),
                    NumberStyles.Integer, CultureInfo.InvariantCulture, out int current)
                    || current < 0 || current <= previous)
                    return false;
                if (current >= frame && insertAt == frames.Count) insertAt = i;
                previous = current;
            }
            if (offset == 1 && frames.Count > 0 && Convert.ToInt32(frames[0], CultureInfo.InvariantCulture) == 0)
                return false; // Do not duplicate the implicit start frame.

            if (offset == 1 && frame == 0) return Replace(values, 0, value);
            if (insertAt < frames.Count
                && Convert.ToInt32(frames[insertAt], CultureInfo.InvariantCulture) == frame)
                return Replace(values, insertAt + offset, value);
            if (frame == 0) return false;
            if (!CreateValue(values[0]!, value, out var entry)
                || !CreateFrame(frames, frame, out var frameEntry))
                return false;

            int valueIndex = insertAt + offset;
            try { values.Insert(valueIndex, entry); }
            catch { return false; }
            bool frameInserted = false;
            try
            {
                frames.Insert(insertAt, frameEntry);
                frameInserted = true;
                if (frames.Count + offset == values.Count
                    && Convert.ToInt32(frames[insertAt], CultureInfo.InvariantCulture) == frame
                    && Number(values[valueIndex], out var actual)
                    && actual == Convert.ToDouble(ReadValue(entry), CultureInfo.InvariantCulture))
                    return true;
            }
            catch { /* Restore the two collections below. */ }
            if (frameInserted)
            {
                try { frames.RemoveAt(insertAt); } catch { }
            }
            try { values.RemoveAt(valueIndex); } catch { }
            return false;
        }

        private static bool Replace(IList values, int index, double value)
        {
            if (index < 0 || index >= values.Count || values[index] == null) return false;
            var old = values[index]!;
            try
            {
                if (IsNumeric(old.GetType()))
                {
                    var converted = Convert.ChangeType(value, old.GetType(), CultureInfo.InvariantCulture);
                    values[index] = converted;
                    if (Number(values[index], out var actual)
                        && actual == Convert.ToDouble(converted, CultureInfo.InvariantCulture))
                        return true;
                    values[index] = old;
                    return false;
                }
                var property = old.GetType().GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
                if (property?.CanWrite != true || !IsNumeric(property.PropertyType)) return false;
                var previous = property.GetValue(old);
                var convertedValue = Convert.ChangeType(value, property.PropertyType, CultureInfo.InvariantCulture);
                property.SetValue(old, convertedValue);
                if (Number(old, out var current)
                    && current == Convert.ToDouble(convertedValue, CultureInfo.InvariantCulture))
                    return true;
                property.SetValue(old, previous);
            }
            catch { return false; }
            return false;
        }

        private static bool CreateValue(object original, double value, out object? entry)
        {
            entry = null;
            try
            {
                var type = original.GetType();
                if (IsNumeric(type))
                {
                    entry = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
                    return true;
                }
                var property = type.GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
                if (property?.CanWrite != true || !IsNumeric(property.PropertyType)) return false;
                var created = Activator.CreateInstance(type);
                if (created == null) return false;
                property.SetValue(created, Convert.ChangeType(value, property.PropertyType, CultureInfo.InvariantCulture));
                entry = created;
                return true;
            }
            catch { return false; }
        }

        private static bool CreateFrame(IList frames, int frame, out object? entry)
        {
            entry = null;
            var type = frames.Count > 0 && frames[0] != null ? frames[0]!.GetType()
                : frames.GetType().GetInterfaces()
                    .FirstOrDefault(i => i.IsGenericType
                        && i.GetGenericTypeDefinition() == typeof(System.Collections.Generic.IList<>))
                    ?.GetGenericArguments()[0] ?? typeof(int);
            try
            {
                if (!IsNumeric(type)) return false;
                entry = Convert.ChangeType(frame, type, CultureInfo.InvariantCulture);
                return true;
            }
            catch { return false; }
        }

        private static bool Number(object? entry, out double value)
        {
            value = 0;
            var raw = ReadValue(entry);
            return raw != null && double.TryParse(Convert.ToString(raw, CultureInfo.InvariantCulture),
                NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value);
        }

        private static object? ReadValue(object? entry) =>
            entry == null ? null : IsNumeric(entry.GetType()) ? entry : Read(entry, "Value");

        private static bool IsNumeric(Type type) => Type.GetTypeCode(type) is
            TypeCode.Byte or TypeCode.SByte or TypeCode.Int16 or TypeCode.UInt16
            or TypeCode.Int32 or TypeCode.UInt32 or TypeCode.Int64 or TypeCode.UInt64
            or TypeCode.Single or TypeCode.Double or TypeCode.Decimal;

        private static object? Read(object source, string name)
        {
            try { return source.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance)?.GetValue(source); }
            catch { return null; }
        }
    }
}
