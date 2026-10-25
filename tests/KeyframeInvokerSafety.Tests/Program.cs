using System;
using System.Threading.Tasks;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

// A boxed int is already assignable to object; Convert.ChangeType(int, object)
// used to reject it, even though the native overload was compatible.
var boxed = new ObjectParameters();
Check(KeyframeMethodInvoker.TryInvokeNamed(boxed, new[] { "AddKeyFrame" }, 12, 2.5),
    "object parameter overload was skipped");
Check(boxed.Calls == 1, "object parameter method not called exactly once");

// Reflection accepts boxed T for Nullable<T>. Keep nullable YMM4 overloads usable.
var nullable = new NullableParameters();
Check(KeyframeMethodInvoker.TryInvokeNamed(nullable, new[] { "AddKeyFrame" }, 9, 1.25),
    "nullable parameter overload was skipped");
Check(nullable.Calls == 1, "nullable overload not called exactly once");

// A non-integral frame must not be rounded to a different timeline position.
var integer = new IntegerFrame();
Check(!KeyframeMethodInvoker.TryInvokeNamed(integer, new[] { "AddKeyFrame" }, 1.5, 1.0),
    "fractional frame was rounded and applied");
Check(integer.Calls == 0, "fractional frame caused a mutation");
Check(KeyframeMethodInvoker.TryInvokeNamed(integer, new[] { "AddKeyFrame" }, 2.0, 1.0),
    "integral double frame should remain compatible");
Check(integer.LastFrame == 2, "integral frame changed during conversion");
Check(!KeyframeMethodInvoker.TryInvokeNamed(integer, new[] { "AddKeyFrame" }, double.PositiveInfinity, 1.0),
    "infinite frame was accepted");
Check(integer.Calls == 1, "invalid frame caused an extra mutation");
Check(!KeyframeMethodInvoker.TryInvokeNamed(integer, new[] { "AddKeyFrame" }, 2, double.NaN),
    "directly assignable NaN was accepted");
Check(integer.Calls == 1, "NaN value caused an extra mutation");

// A conversion to float can overflow to Infinity without throwing.
var floating = new FloatValue();
Check(!KeyframeMethodInvoker.TryInvokeNamed(floating, new[] { "AddKeyFrame" }, 1, 1e100),
    "overflowed float value was accepted");
Check(floating.Calls == 0, "overflowed value caused a mutation");

// Returning a Task is not proof that a native edit succeeded. Skip async
// overloads rather than reporting a pending/failed operation as successful.
var asynchronous = new AsyncParameters();
Check(!KeyframeMethodInvoker.TryInvokeNamed(asynchronous, new[] { "AddKeyFrame" }, 3, 1.0),
    "pending Task was reported as a successful edit");
Check(asynchronous.Calls == 0, "unsupported async overload was invoked");
var valueTask = new ValueTaskParameters();
Check(!KeyframeMethodInvoker.TryInvokeNamed(valueTask, new[] { "AddKeyFrame" }, 3, 1.0),
    "pending ValueTask was reported as a successful edit");
Check(valueTask.Calls == 0, "unsupported ValueTask overload was invoked");

// Existing behavior: optional arguments, bool refusal and later-name fallback.
var fallback = new Fallback();
Check(KeyframeMethodInvoker.TryInvokeNamed(fallback, new[] { "AddKeyFrame", "SetKeyFrame" }, 4, 1.0),
    "fallback after bool false stopped working");
Check(fallback.Refused == 1 && fallback.Accepted == 1, "fallback invoked unexpected methods");
var optional = new Optional();
Check(KeyframeMethodInvoker.TryInvokeNamed(optional, new[] { "AddKeyFrame" }, 4),
    "optional parameter stopped working");
Check(optional.Calls == 1, "optional overload not called");
Console.WriteLine("Keyframe invocation safety tests passed");

internal sealed class ObjectParameters
{
    public int Calls;
    public bool AddKeyFrame(object frame, object value)
    {
        Calls++;
        return frame is int i && i == 12 && value is double d && d == 2.5;
    }
}
internal sealed class NullableParameters
{
    public int Calls;
    public bool AddKeyFrame(int? frame, double? value)
    {
        Calls++;
        return frame == 9 && value == 1.25;
    }
}
internal sealed class IntegerFrame
{
    public int Calls;
    public int LastFrame;
    public bool AddKeyFrame(int frame, double value)
    {
        Calls++;
        LastFrame = frame;
        return true;
    }
}
internal sealed class FloatValue
{
    public int Calls;
    public bool AddKeyFrame(int frame, float value) { Calls++; return true; }
}
internal sealed class AsyncParameters
{
    public int Calls;
    public Task<bool> AddKeyFrame(int frame, double value)
    {
        Calls++;
        return Task.FromResult(false);
    }
}
internal sealed class ValueTaskParameters
{
    public int Calls;
    public ValueTask<bool> AddKeyFrame(int frame, double value)
    {
        Calls++;
        return ValueTask.FromResult(false);
    }
}
internal sealed class Fallback
{
    public int Refused;
    public int Accepted;
    public bool AddKeyFrame(int frame, double value) { Refused++; return false; }
    public bool SetKeyFrame(int frame, double value) { Accepted++; return true; }
}
internal sealed class Optional
{
    public int Calls;
    public void AddKeyFrame(int frame, double value = 0) { Calls++; }
}
