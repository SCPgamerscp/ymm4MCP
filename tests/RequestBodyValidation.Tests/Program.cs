using System;
using System.Text;
using System.Text.Json;
using YMM4McpPlugin;

static Dictionary<string, JsonElement> Parse(string json) =>
    RequestBodyValidation.Parse(Encoding.UTF8.GetBytes(json));

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

static void Reject<T>(Action action, string message) where T : Exception
{
    try { action(); }
    catch (T) { return; }
    throw new Exception(message);
}

// Normal clients, including Unicode payloads and empty legacy POSTs, still work.
var valid = Parse("{\"name\":\"霊夢\",\"frame\":10,\"items\":[{\"id\":\"one\"}]}");
Check(valid["name"].GetString() == "霊夢" && valid["frame"].GetInt32() == 10, "valid JSON was changed");
Check(valid["items"][0].GetProperty("id").GetString() == "one", "nested JSON was changed");
Check(RequestBodyValidation.Parse(ReadOnlyMemory<byte>.Empty).Count == 0, "empty body compatibility");

// Duplicate keys at any depth are ambiguous edit instructions and must fail.
Reject<ArgumentException>(() => Parse("{\"frame\":1,\"frame\":2}"), "duplicate top-level key accepted");
Reject<ArgumentException>(() => Parse("{\"edit\":{\"layer\":1,\"layer\":2}}"), "nested duplicate accepted");
Reject<ArgumentException>(() => Parse("{\"edits\":[{\"at\":1,\"at\":2}]}"), "array-nested duplicate accepted");
Reject<ArgumentException>(() => Parse("{\"a\":1,\"\\u0061\":2}"), "escaped duplicate accepted");
Check(Parse("{\"a\":1,\"A\":2}").Count == 2, "distinct case-sensitive keys rejected");

// Reject malformed byte sequences rather than accepting replacement characters.
byte[] invalidUtf8 = { (byte)'{', (byte)'"', (byte)'x', (byte)'"', (byte)':',
                       (byte)'"', 0xC3, 0x28, (byte)'"', (byte)'}' };
Reject<ArgumentException>(() => RequestBodyValidation.Parse(invalidUtf8), "invalid UTF-8 accepted");

// A JSON scalar or array cannot be treated as an empty command dictionary.
foreach (string invalidRoot in new[] { "null", "[]", "true", "123", "\"text\"" })
    Reject<ArgumentException>(() => Parse(invalidRoot), "non-object JSON root accepted: " + invalidRoot);
Reject<JsonException>(() => Parse("{\"frame\":"), "truncated JSON accepted");
Reject<JsonException>(() => Parse("{\"a\":1,}"), "trailing comma accepted");
Reject<JsonException>(() => Parse("{/*comment*/\"a\":1}"), "JSON comments accepted");
string nested = new string('[', 40) + "0" + new string(']', 40);
Reject<JsonException>(() => Parse("{\"x\":" + nested + "}"), "excessive nesting accepted");

// HttpListener returns -1 for chunked requests; declared lengths must match.
RequestBodyValidation.ValidateLength(-1, 12);
RequestBodyValidation.ValidateLength(0, 0);
RequestBodyValidation.ValidateLength(12, 12);
Reject<ArgumentException>(() => RequestBodyValidation.ValidateLength(12, 11),
    "truncated declared body accepted");
Reject<ArgumentException>(() => RequestBodyValidation.ValidateLength(12, 13),
    "body longer than declared length accepted");
Reject<ArgumentException>(() => RequestBodyValidation.ValidateLength(2, 0),
    "missing body accepted with positive declared length");

Console.WriteLine("Native HTTP request body safety tests passed");
