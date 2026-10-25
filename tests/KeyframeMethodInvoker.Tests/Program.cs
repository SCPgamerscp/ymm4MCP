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
var shared = new MockKeyFrames();
var detached = new MockAnimation();
Check(KeyframeHostBinding.TryBind(new MockItem(shared), detached),
    "detached Animation must bind to the item's shared KeyFrames");
Check(ReferenceEquals(detached.KeyFrames, shared) && detached.Binds == 1,
    "binding must use the exact shared KeyFrames instance");
Check(KeyframeHostBinding.TryBind(new MockItem(shared), detached),
    "already-bound Animation should be accepted when it owns the same shared frames");
Check(!KeyframeHostBinding.HasConflictingBinding(new MockItem(shared), detached),
    "matching item and animation frames must not be treated as a conflict");
Check(ReferenceEquals(detached.KeyFrames, shared) && detached.Binds == 1,
    "existing binding must not be overwritten");
var foreignFrames = new MockKeyFrames();
Check(KeyframeHostBinding.HasConflictingBinding(new MockItem(foreignFrames), detached),
    "cross-item keyframe binding must be detected before editing");
Check(!KeyframeHostBinding.TryBind(new MockItem(foreignFrames), detached),
    "cross-item binding must not be reported as successful");
Check(ReferenceEquals(detached.KeyFrames, shared) && detached.Binds == 1,
    "cross-item binding check must not mutate the original animation");
Check(!KeyframeHostBinding.HasConflictingBinding(new MockItem(null), detached),
    "hosts without item-level KeyFrames should retain their legacy fallback");
Check(!KeyframeHostBinding.TryBind(new MockItem(null), new MockAnimation()),
    "missing item KeyFrames must not be fabricated");
Check(!KeyframeHostBinding.TryBind(new MockItem(shared), new WrongBindingAnimation()),
    "incompatible SetKeyFrames signature must be rejected");
Check(!KeyframeHostBinding.TryBind(new MockItem(shared), new ThrowingBindingAnimation()),
    "host SetKeyFrames failures must not be reported as success");
Check(!KeyframeHostBinding.TryBind(new MockItem(shared), new NoOpBindingAnimation()),
    "a no-op SetKeyFrames must not be reported as success");

var overloaded = new OverloadedBindingAnimation();
Check(KeyframeHostBinding.TryBind(new MockItem(shared), overloaded),
    "prefer the strongly typed SetKeyFrames overload over a broad object overload");
Check(ReferenceEquals(overloaded.KeyFrames, shared) && overloaded.BroadCalls == 0,
    "a broad overload must not shadow the strongly typed overload");

var fallbackOverload = new FallbackBindingAnimation();
Check(KeyframeHostBinding.TryBind(new MockItem(shared), fallbackOverload),
    "try a compatible fallback after a throwing SetKeyFrames overload");
Check(ReferenceEquals(fallbackOverload.KeyFrames, shared) && fallbackOverload.TypedCalls == 1,
    "failed typed overload must allow fallback without replacing the collection");

var noOpOverload = new NoOpOverloadBindingAnimation();
Check(KeyframeHostBinding.TryBind(new MockItem(shared), noOpOverload),
    "try another compatible overload after a silent no-op");
Check(ReferenceEquals(noOpOverload.KeyFrames, shared) && noOpOverload.TypedCalls == 1,
    "no-op typed overload must not be treated as success");


Console.WriteLine("Keyframe method invocation and host binding tests passed");

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

internal sealed class MockKeyFrames
{
    public List<int> Frames { get; } = new();
}

internal sealed class MockItem(MockKeyFrames? frames)
{
    public MockKeyFrames? KeyFrames { get; } = frames;
}

internal sealed class MockAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public int Binds { get; private set; }
    public void SetKeyFrames(MockKeyFrames frames) { KeyFrames = frames; Binds++; }
}

internal sealed class WrongBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public void SetKeyFrames(string frames) { }
}

internal sealed class ThrowingBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public void SetKeyFrames(MockKeyFrames frames) => throw new InvalidOperationException("host refused");
}

internal sealed class NoOpBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public void SetKeyFrames(MockKeyFrames frames) { }
}

internal sealed class OverloadedBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public int BroadCalls { get; private set; }
    public void SetKeyFrames(object frames) { BroadCalls++; throw new NotSupportedException(); }
    public void SetKeyFrames(MockKeyFrames frames) => KeyFrames = frames;
}

internal sealed class FallbackBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public int TypedCalls { get; private set; }
    public void SetKeyFrames(MockKeyFrames frames) { TypedCalls++; throw new NotSupportedException(); }
    public void SetKeyFrames(object frames) => KeyFrames = (MockKeyFrames)frames;
}

internal sealed class NoOpOverloadBindingAnimation
{
    public MockKeyFrames? KeyFrames { get; private set; }
    public int TypedCalls { get; private set; }
    public void SetKeyFrames(MockKeyFrames frames) { TypedCalls++; }
    public void SetKeyFrames(object frames) => KeyFrames = (MockKeyFrames)frames;
}
