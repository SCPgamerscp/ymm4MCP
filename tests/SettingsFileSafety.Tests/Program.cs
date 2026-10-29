using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YMM4McpPlugin;

static void Check(bool ok, string message)
{
    if (!ok) throw new Exception(message);
}

static void ExpectInvalid(string path, string contents)
{
    File.WriteAllText(path, contents);
    try
    {
        SettingsFileSafety.ReadValidated(path);
        throw new Exception("Malformed settings were accepted: " + contents);
    }
    catch (InvalidDataException) { }
}

string root = Path.Combine(Path.GetTempPath(), "ymm4-settings-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    string path = Path.Combine(root, "settings.json");
    File.WriteAllText(path, "{\"Port\":8765,\"AutoStart\":true}");
    Check(SettingsFileSafety.ReadValidated(path).Contains("\"Port\":8765"),
        "ordinary settings must remain readable");

    // A duplicate key must not silently change listener or reflection settings.
    ExpectInvalid(path, "{\"Port\":8765,\"Port\":9999}");
    ExpectInvalid(path, "{\"AutoStart\":false,\"AutoStart\":true}");
    ExpectInvalid(path, "{\"nested\":{\"token\":\"first\",\"token\":\"second\"}}");
    ExpectInvalid(path, "{\"items\":[{\"key\":1,\"key\":2}]}");
    ExpectInvalid(path, "{\"AutoStrat\":true}");
    ExpectInvalid(path, "{\"port\":8765}");
    ExpectInvalid(path, "{\"Port\":8765,\"UnexpectedFlag\":true}");

    // A 64KiB cap must apply to actual bytes, not a stale FileInfo length.
    File.WriteAllText(path, "{\"padding\":\"" + new string('x', 65536) + "\"}");
    try
    {
        SettingsFileSafety.ReadValidated(path);
        throw new Exception("oversized settings accepted");
    }
    catch (InvalidDataException) { }

    // Strict UTF-8 decoding prevents invalid byte sequences from being normalized.
    File.WriteAllBytes(path, new byte[] { (byte)'{', (byte)'"', 0xff, (byte)'"', (byte)':', (byte)'1', (byte)'}' });
    try
    {
        SettingsFileSafety.ReadValidated(path);
        throw new Exception("invalid UTF-8 accepted");
    }
    catch (InvalidDataException ex)
    {
        Check(ex.InnerException is DecoderFallbackException,
            "malformed UTF-8 must be normalized to InvalidDataException for McpSettings.Load");
    }

    File.WriteAllText(path, "{\"Port\":8765}", new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    Check(SettingsFileSafety.ReadValidated(path).Contains("\"Port\":8765"),
        "UTF-8 BOM settings must remain compatible");

    SettingsFileSafety.WriteAtomic(path, "{\"Port\":8766}");
    Check(SettingsFileSafety.ReadValidated(path) == "{\"Port\":8766}",
        "atomic save must replace the previous file");

    // Two simultaneous writers must not race over a fixed staging filename.
    var writers = Enumerable.Range(0, 12).Select(i =>
        Task.Run(() => SettingsFileSafety.WriteAtomic(path, "{\"Port\":" + (9000 + i) + "}")));
    await Task.WhenAll(writers);
    Check(SettingsFileSafety.ReadValidated(path).Contains("\"Port\":9"),
        "concurrent saves must leave a complete JSON document");
    Check(Directory.GetFiles(root, "settings.json.tmp.*").Length == 0,
        "successful saves must not leave staging files");

    string missing = Path.Combine(root, "missing", "settings.json");
    try
    {
        SettingsFileSafety.WriteAtomic(missing, "{}");
        throw new Exception("save to nonexistent directory succeeded");
    }
    catch (DirectoryNotFoundException) { }
    Check(!Directory.Exists(Path.GetDirectoryName(missing)!),
        "failed save must not create unexpected directories");
}
finally
{
    Directory.Delete(root, recursive: true);
}
Console.WriteLine("Settings file safety regression tests passed");
