using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var inactive = new MainViewModel { Index = 0 };
var active = new MainViewModel { Index = 3 };
var candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (new OtherViewModel { Index = 0 }, true, true),
    (inactive, false, true),
    (active, true, true)
});
Check(ReferenceEquals(candidate, active), "Active editor with nonzero Index was not selected");

candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (new OtherViewModel { Index = 0 }, true, true),
    (inactive, false, true),
    (active, false, false)
});
Check(ReferenceEquals(candidate, inactive), "Visible editor fallback failed");
Check(MainViewModelSelector.Select(new (object?, bool, bool)[]
    { (new OtherViewModel { Index = 0 }, true, true) }) == null,
    "Unrelated active view model was selected");

// When no window is active, picking the first of two visible editors would
// silently edit the wrong project. Require a unique highest-priority editor.
var secondVisible = new MainViewModel { Index = 7 };
candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (inactive, false, true),
    (secondVisible, false, true)
});
Check(candidate == null, "Ambiguous visible editors should not be selected");

// Two windows claiming to be active are equally unsafe.
candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (active, true, true),
    (secondVisible, true, true)
});
Check(candidate == null, "Ambiguous active editors should not be selected");

// Duplicate observations of the SAME editor should not become ambiguous.
candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (active, true, true),
    (active, true, true),
    (inactive, false, true)
});
Check(ReferenceEquals(candidate, active), "Repeated same editor was rejected");

// A unique hidden editor remains available as a backwards-compatible fallback.
candidate = MainViewModelSelector.Select(new (object?, bool, bool)[]
{
    (inactive, false, false)
});
Check(ReferenceEquals(candidate, inactive), "Unique hidden editor fallback failed");


// A window closing during enumeration must not turn a partial snapshot
// into a successful selection of the wrong editor.
static IEnumerable<(object?, bool, bool)> ChangingWindows()
{
    yield return (new MainViewModel { Index = 8 }, true, true);
    throw new InvalidOperationException("collection modified");
}
Check(MainViewModelSelector.Select(ChangingWindows()) == null,
    "Partially enumerated editor list should fail closed");

Console.WriteLine("MainViewModel selection tests passed");

class MainViewModel { public int Index { get; set; } }
class OtherViewModel { public int Index { get; set; } }
