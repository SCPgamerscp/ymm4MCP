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
                return TryUpsertImmutable(animation, owner, frames, values, frame, value);

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
            // Some third-party IList implementations insert at a different index
            // without throwing. Snapshot the original order before mutating so a
            // false success cannot attach a keyframe to the wrong frame.
            object?[] originalValues;
            object?[] originalFrames;
            try
            {
                originalValues = values.Cast<object?>().ToArray();
                originalFrames = frames.Cast<object?>().ToArray();
            }
            catch { return false; }

            try { values.Insert(valueIndex, entry); }
            catch
            {
                RestoreSingleInsert(values, originalValues);
                return false;
            }
            if (values.Count != originalValues.Length + 1)
                return false; // A silent no-op is not a successful insertion.
            try
            {
                frames.Insert(insertAt, frameEntry);
                if (InsertedAt(frames, originalFrames, insertAt, frameEntry)
                    && InsertedAt(values, originalValues, valueIndex, entry)
                    && frames.Count + offset == values.Count
                    && Convert.ToInt32(frames[insertAt], CultureInfo.InvariantCulture) == frame
                    && Number(values[valueIndex], out var actual)
                    && actual == Convert.ToDouble(ReadValue(entry), CultureInfo.InvariantCulture))
                    return true;
            }
            catch { /* Restore the two collections below. */ }

            // Do not blindly RemoveAt(requestedIndex): a non-conforming IList
            // may have inserted elsewhere, which would delete an original value.
            RestoreSingleInsert(frames, originalFrames);
            RestoreSingleInsert(values, originalValues);
            return false;
        }

        private static bool InsertedAt(IList list, object?[] before, int at, object? added)
        {
            if (at < 0 || at > before.Length || list.Count != before.Length + 1)
                return false;
            try
            {
                for (int i = 0; i < list.Count; i++)
                {
                    object? expected = i == at ? added : before[i < at ? i : i - 1];
                    if (!Equals(list[i], expected)) return false;
                }
                return true;
            }
            catch { return false; }
        }

        // Restore only when removing exactly one element reproduces the
        // original sequence. This also handles insert-then-throw at a wrong
        // index without deleting a pre-existing frame or value.
        private static void RestoreSingleInsert(IList list, object?[] before)
        {
            try
            {
                if (list.Count != before.Length + 1) return;
                for (int removed = 0; removed < list.Count; removed++)
                {
                    bool matches = true;
                    for (int i = 0; i < before.Length; i++)
                    {
                        if (!Equals(list[i < removed ? i : i + 1], before[i]))
                        {
                            matches = false;
                            break;
                        }
                    }
                    if (!matches) continue;
                    list.RemoveAt(removed);
                    return;
                }
            }
            catch { /* Fail closed; the caller reports failure. */ }
        }


        // ImmutableList<T> implements IList but rejects in-place Insert/SetItem.
        // Only replace collections when both owners expose public writable properties.
        // Mixed mutable/immutable layouts are rejected to avoid half-applied changes.
        private static bool TryUpsertImmutable(
            object animation, object owner, IList frames, IList values, int frame, double value)
        {
            if (!(frames.IsReadOnly || frames.IsFixedSize)
                || !(values.IsReadOnly || values.IsFixedSize))
                return false;
            var frameProperty = owner.GetType().GetProperty("Frames", BindingFlags.Public | BindingFlags.Instance)
                ?? owner.GetType().GetProperty("FrameList", BindingFlags.Public | BindingFlags.Instance);
            var valueProperty = animation.GetType().GetProperty("Values", BindingFlags.Public | BindingFlags.Instance);
            if (frameProperty?.SetMethod?.IsPublic != true
                || valueProperty?.SetMethod?.IsPublic != true)
                return false;
            try
            {
                if (!ReferenceEquals(frameProperty.GetValue(owner), frames)
                    || !ReferenceEquals(valueProperty.GetValue(animation), values))
                    return false;
            }
            catch
            {
                // A YMM4 property getter can throw while the editor is closing.
                // Treat it as an unsupported layout, not an uncaught API error.
                return false;
            }

            int offset = values.Count == frames.Count + 1 ? 1
                : values.Count == frames.Count && frames.Count > 0 ? 0 : -1;
            if (offset < 0 || values.Count == 0) return false;

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
            if (offset == 1 && frames.Count > 0
                && Convert.ToInt32(frames[0], CultureInfo.InvariantCulture) == 0)
                return false;

            int valueIndex;
            bool insert = false;
            if (offset == 1 && frame == 0) valueIndex = 0;
            else if (insertAt < frames.Count
                && Convert.ToInt32(frames[insertAt], CultureInfo.InvariantCulture) == frame)
                valueIndex = insertAt + offset;
            else
            {
                if (frame == 0) return false;
                valueIndex = insertAt + offset;
                insert = true;
            }

            // Only numeric values are supported here: cloning arbitrary animation
            // objects could silently discard easing and interpolation metadata.
            if (values[0] == null || !IsNumeric(values[0]!.GetType()))
                return false;
            if (!CreateValue(values[0]!, value, out var entry)) return false;
            // A repeated request for the same value is already satisfied.
            // ImmutableList<T>.SetItem may return the original instance for no-op updates.
            if (!insert && Number(values[valueIndex], out var existing)
                && existing == Convert.ToDouble(entry, CultureInfo.InvariantCulture))
                return true;
            object? nextFrames = frames;
            object? nextValues;
            if (insert)
            {
                if (!CreateFrame(frames, frame, out var frameEntry)
                    || !TryImmutableCall(frames, "Insert", insertAt, frameEntry, out nextFrames)
                    || !TryImmutableCall(values, "Insert", valueIndex, entry, out nextValues))
                    return false;
            }
            else if (!TryImmutableCall(values, "SetItem", valueIndex, entry, out nextValues))
                return false;

            if (nextFrames is not IList newFrames || nextValues is not IList newValues
                || newFrames.Count + offset != newValues.Count
                || !frameProperty.PropertyType.IsInstanceOfType(nextFrames)
                || !valueProperty.PropertyType.IsInstanceOfType(nextValues)
                || !Number(newValues[valueIndex], out var actual)
                || actual != Convert.ToDouble(entry, CultureInfo.InvariantCulture))
                return false;

            // A reflected immutable Insert method is not guaranteed to insert the
            // requested frame. Verify the entire ordered frame sequence before
            // committing either collection, not just the resulting counts.
            if (insert)
            {
                for (int i = 0; i < newFrames.Count; i++)
                {
                    var expected = i == insertAt ? frame : frames[i < insertAt ? i : i - 1];
                    if (!int.TryParse(Convert.ToString(expected, CultureInfo.InvariantCulture),
                            NumberStyles.Integer, CultureInfo.InvariantCulture, out int expectedFrame)
                        || !int.TryParse(Convert.ToString(newFrames[i], CultureInfo.InvariantCulture),
                            NumberStyles.Integer, CultureInfo.InvariantCulture, out int actualFrame)
                        || actualFrame != expectedFrame)
                        return false;
                }
            }

            // A reflected Insert/SetItem may return a collection of the expected size
            // while silently altering unrelated keyframe values. Check every element
            // before replacing either collection, including the requested value.
            if (newValues.Count != values.Count + (insert ? 1 : 0))
                return false;
            for (int i = 0; i < newValues.Count; i++)
            {
                var expected = insert
                    ? (i == valueIndex ? entry : values[i < valueIndex ? i : i - 1])
                    : (i == valueIndex ? entry : values[i]);
                if (!Equals(newValues[i], expected))
                    return false;
            }

            // Stage both new immutable lists before touching the animation.
            // Restore both references if either property setter rejects the update.
            try
            {
                valueProperty.SetValue(animation, nextValues);
                frameProperty.SetValue(owner, nextFrames);
                if (ReferenceEquals(valueProperty.GetValue(animation), nextValues)
                    && ReferenceEquals(frameProperty.GetValue(owner), nextFrames))
                    return true;
            }
            catch { /* Best-effort restoration below. */ }
            try { frameProperty.SetValue(owner, frames); } catch { }
            try { valueProperty.SetValue(animation, values); } catch { }
            return false;
        }

        private static bool TryImmutableCall(
            object source, string name, int index, object? item, out object? updated)
        {
            updated = null;
            try
            {
                var method = source.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == name && m.GetParameters().Length == 2
                        && m.GetParameters()[0].ParameterType == typeof(int)
                        && item != null && m.GetParameters()[1].ParameterType.IsInstanceOfType(item));
                if (method == null) return false;
                updated = method.Invoke(source, new[] { (object)index, item });
                return updated != null && !ReferenceEquals(updated, source);
            }
            catch { return false; }
        }

        private static bool Replace(IList values, int index, double value)
        {
            if (index < 0 || index >= values.Count || values[index] == null) return false;
            var old = values[index]!;
            if (IsNumeric(old.GetType()))
            {
                if (!TryConvertNumeric(value, old.GetType(), out var converted)) return false;
                try
                {
                    values[index] = converted;
                    if (Number(values[index], out var actual)
                        && actual == Convert.ToDouble(converted, CultureInfo.InvariantCulture))
                        return true;
                }
                catch { /* The IList may have changed before throwing. */ }

                try
                {
                    if (Equals(values[index], old)) return false;
                    values[index] = old;
                    if (Equals(values[index], old)) return false;
                }
                catch { /* Restoration must be verified below. */ }
                throw new InvalidOperationException(
                    "Keyframe replacement failed and the original numeric value could not be restored.");
            }

            var property = old.GetType().GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            if (property?.CanWrite != true || !IsNumeric(property.PropertyType)) return false;
            object? previous;
            try { previous = property.GetValue(old); }
            catch { return false; }
            if (!TryConvertNumeric(value, property.PropertyType, out var convertedValue)) return false;
            try
            {
                property.SetValue(old, convertedValue);
                if (Number(old, out var actual)
                    && actual == Convert.ToDouble(convertedValue, CultureInfo.InvariantCulture))
                    return true;
            }
            catch { /* A reflected setter can mutate before throwing. */ }

            try
            {
                if (Equals(property.GetValue(old), previous)) return false;
                property.SetValue(old, previous);
                if (Equals(property.GetValue(old), previous)) return false;
            }
            catch { /* Do not continue editing an unverified value. */ }
            throw new InvalidOperationException(
                "Keyframe replacement failed and the original Value property could not be restored.");
        }

        private static bool CreateValue(object original, double value, out object? entry)
        {
            entry = null;
            try
            {
                var type = original.GetType();
                if (IsNumeric(type))
                {
                    return TryConvertNumeric(value, type, out entry);
                }
                var property = type.GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
                if (property?.CanWrite != true || !IsNumeric(property.PropertyType)) return false;
                var created = Activator.CreateInstance(type);
                if (created == null) return false;
                if (!TryConvertNumeric(value, property.PropertyType, out var converted)) return false;
                property.SetValue(created, converted);
                entry = created;
                return true;
            }
            catch { return false; }
        }

        // Convert.ChangeType rounds fractional doubles when the target is an
        // integer and can overflow float to Infinity without throwing.
        // Neither outcome represents the value the caller requested.
        private static bool TryConvertNumeric(double value, Type type, out object? converted)
        {
            converted = null;
            if (!double.IsFinite(value) || !IsNumeric(type)) return false;
            if (type != typeof(float) && type != typeof(double) && type != typeof(decimal)
                && value != Math.Truncate(value))
                return false;
            try
            {
                var result = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
                if (result is float f && !float.IsFinite(f)) return false;
                if (result is double d && !double.IsFinite(d)) return false;
                converted = result;
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
