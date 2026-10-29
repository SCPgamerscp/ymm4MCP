using System.Reflection;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var partial = new ReflectionTypeLoadException(
    new Type?[] { typeof(string), null, typeof(int) },
    new Exception[] { new TypeLoadException("unavailable") });
var recovered = DebugInspectionSafety.RecoverLoadableTypes(partial);
Check(recovered.SequenceEqual(new[] { typeof(string), typeof(int) }),
      "partially loadable types were discarded");
Check(DebugInspectionSafety.LoadableTypes(typeof(Program).Assembly).Length > 0,
      "healthy assembly returned no types");

int inspected = 0;
IEnumerable<object?> Endless()
{
    while (true)
    {
        inspected++;
        yield return new MenuEntry();
    }
}
var menu = DebugInspectionSafety.SnapshotMenuItems(Endless(), maxItems: 5);
Check(menu.Length == 5 && inspected == 5, "menu enumeration exceeded its bound");
Check(DebugInspectionSafety.SnapshotMenuItems(null).Length == 0, "null menu was not handled");
Check(DebugInspectionSafety.SnapshotMenuItems(new object?[] { null, new MenuEntry() }).Length == 1,
      "null menu entry was not skipped");
Check(DebugInspectionSafety.SnapshotMenuItems(new object[] { new MenuEntry() }, 0).Length == 1,
      "zero limit was not clamped");
Console.WriteLine("Debug inspection safety tests passed");

public sealed class MenuEntry
{
    public string Name => throw new InvalidOperationException("getter must not execute");
}
