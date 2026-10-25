using System.Reflection;
using System.Text.Json;
using YMM4McpPlugin;

static JsonElement Json(string raw) => JsonDocument.Parse(raw).RootElement.Clone();
static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

Check(ReflectionInput.TryConvert(Json("9223372036854775807"), typeof(long), out var maxLong)
      && maxLong is long l && l == long.MaxValue, "long lost precision");
Check(ReflectionInput.TryConvert(Json("18446744073709551615"), typeof(ulong), out var maxUlong)
      && maxUlong is ulong ul && ul == ulong.MaxValue, "ulong lost precision");
Check(!ReflectionInput.TryConvert(Json("9223372036854775808"), typeof(long), out _),
      "overflowing long was accepted");
Check(ReflectionInput.TryConvert(Json("1e3"), typeof(int), out var exponent)
      && exponent is int n && n == 1000, "integral exponent was rejected");
Check(!ReflectionInput.TryConvert(Json("1.5"), typeof(int), out _),
      "fractional number was rounded to integer");
Check(!ReflectionInput.TryConvert(Json("1e400"), typeof(double), out _),
      "infinite double was accepted");
Check(!ReflectionInput.TryConvert(Json("\"Infinity\""), typeof(double), out _),
      "infinite double string was accepted");
Check(!ReflectionInput.TryConvert(Json("{\"value\":1}"), typeof(int), out _),
      "object silently became a numeric string");
Check(!ReflectionInput.TryConvert(Json("true"), typeof(int), out _),
      "boolean silently became an integer");
Check(!ReflectionInput.TryConvert(Json("null"), typeof(int), out _),
      "null became a non-nullable value type");
Check(ReflectionInput.TryConvert(Json("null"), typeof(int?), out var nullable)
      && nullable == null, "nullable null rejected");
Check(ReflectionInput.TryConvert(Json("\"Friday\""), typeof(DayOfWeek), out var day)
      && day is DayOfWeek.Friday, "enum name not converted");
Check(!ReflectionInput.TryConvert(Json("\"nonsense\""), typeof(DayOfWeek), out _),
      "invalid enum accepted");
Check(ReflectionInput.TryConvert(Json("{\"x\":[1]}"), typeof(JsonElement), out var element)
      && element is JsonElement je && je.GetProperty("x")[0].GetInt32() == 1,
      "JsonElement did not retain structured data");

var methods = typeof(Target).GetMethods(BindingFlags.Public | BindingFlags.Instance);
Check(ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "Apply"),
      new[] { Json("7") }, out var selected, out var args, out var ambiguous)
      && selected!.GetParameters()[0].ParameterType == typeof(int)
      && args![0] is int i && i == 7 && !ambiguous,
      "numeric overload selection failed");
Check(ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "Apply"),
      new[] { Json("\"hello\"") }, out selected, out args, out ambiguous)
      && selected!.GetParameters()[0].ParameterType == typeof(string),
      "string overload selection failed");
Check(!ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "Ambiguous"),
      new[] { Json("7") }, out selected, out args, out ambiguous)
      && ambiguous && selected == null && args == null,
      "ambiguous overload silently selected");
Check(!ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "OnlyInteger"),
      new[] { Json("\"invalid\"") }, out _, out _, out ambiguous)
      && !ambiguous, "invalid argument accepted");
Check(!ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "Generic"),
      new[] { Json("1") }, out _, out _, out _),
      "open generic method was accepted");
Check(!ReflectionInput.TrySelectMethod(methods.Where(m => m.Name == "ByReference"),
      new[] { Json("1") }, out _, out _, out _),
      "by-reference parameter was accepted");

Console.WriteLine("Reflection input safety tests passed");

public sealed class Target
{
    public void Apply(int value) { }
    public void Apply(string value) { }
    public void Ambiguous(int value) { }
    public void Ambiguous(long value) { }
    public void OnlyInteger(int value) { }
    public void Generic<T>(T value) { }
    public void ByReference(ref int value) { }
}
