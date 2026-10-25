using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var rejects = new RejectingAnimation();
Check(!KeyframeMethodInvoker.TryInvokeNamed(rejects, new[] { "AddKeyFrame" }, 10, 1.5),
    "bool false must not be reported as successful");
Check(rejects.Attempts == 1, "rejecting method should have been invoked");

var fallback = new FallbackAnimation();
Check(KeyframeMethodInvoker.TryInvokeNamed(fallback, new[] { "AddKeyFrame", "SetKeyFrame" }, 10, 1.5),
    "false return should allow a later method name");
Check(fallback.Rejected == 1 && fallback.Applied == 1, "fallback method should run once");

var accepts = new AcceptingAnimation();
Check(KeyframeMethodInvoker.TryInvokeNamed(accepts, new[] { "AddKeyFrame" }, 10, 1.5),
    "bool true must be accepted");
Check(accepts.Applied == 1, "accepted method should run once");

var nullable = new NullableAnimation();
Check(!KeyframeMethodInvoker.TryInvokeNamed(nullable, new[] { "AddKeyFrame" }, 10, 1.5),
    "nullable bool null must not count as success");

var legacy = new LegacyAnimation();
Check(KeyframeMethodInvoker.TryInvokeNamed(legacy, new[] { "AddKeyFrame" }, 10),
    "void methods with optional parameters must remain supported");
Check(legacy.Applied == 1, "legacy void method should run");

var throws = new ThrowingAnimation();
Check(KeyframeMethodInvoker.TryInvokeNamed(throws, new[] { "AddKeyFrame", "SetKeyFrame" }, 10, 1.5),
    "throwing overload should allow a later method");
Check(throws.Applied == 1, "later method should run after exception");

Check(!KeyframeMethodInvoker.TryInvokeNamed(legacy, new[] { "MissingMethod" }, 10),
    "missing method must fail");
Console.WriteLine("Keyframe method invocation tests passed");

internal sealed class RejectingAnimation
{
    public int Attempts { get; private set; }
    public bool AddKeyFrame(int frame, double value) { Attempts++; return false; }
}

internal sealed class FallbackAnimation
{
    public int Rejected { get; private set; }
    public int Applied { get; private set; }
    public bool AddKeyFrame(int frame, double value) { Rejected++; return false; }
    public bool SetKeyFrame(int frame, double value) { Applied++; return true; }
}

internal sealed class AcceptingAnimation
{
    public int Applied { get; private set; }
    public bool AddKeyFrame(int frame, double value) { Applied++; return true; }
}

internal sealed class NullableAnimation
{
    public bool? AddKeyFrame(int frame, double value) => null;
}

internal sealed class LegacyAnimation
{
    public int Applied { get; private set; }
    public void AddKeyFrame(int frame, double value = 0) { Applied++; }
}

internal sealed class ThrowingAnimation
{
    public int Applied { get; private set; }
    public bool AddKeyFrame(int frame, double value) => throw new NotSupportedException();
    public void SetKeyFrame(int frame, double value) { Applied++; }
}
