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
Console.WriteLine("MainViewModel selection tests passed");

class MainViewModel { public int Index { get; set; } }
class OtherViewModel { public int Index { get; set; } }
