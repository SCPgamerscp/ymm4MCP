using YMM4McpPlugin;

static void Check(bool ok, string message)
{
    if (!ok) throw new Exception(message);
}

string root = Path.Combine(Path.GetTempPath(), "ymm4-assets-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    File.WriteAllText(Path.Combine(root, "爆発1.wav"), "same audio bytes");
    File.WriteAllText(Path.Combine(root, "爆発2.wav"), "same audio bytes");
    File.WriteAllText(Path.Combine(root, "notes.txt"), "ignored");
    string child = Path.Combine(root, "sub");
    Directory.CreateDirectory(child);
    File.WriteAllText(Path.Combine(child, "scene.mp4"), "video");
    var flat = AssetScanner.Scan(root, "爆発", hash: true);
    Check(flat.Success && flat.Assets.Length == 2 && flat.Duplicates.Length == 1,
        "matching duplicate media should be returned");
    var all = AssetScanner.Scan(root, recursive: true);
    Check(all.Assets.Length == 3 && all.Assets.Any(a => a.Name == "scene.mp4"), "recursive media missing");
    var limited = AssetScanner.Scan(root, maxResults: 1);
    Check(limited.Assets.Length == 1 && limited.Truncated, "limit should report truncation");
    try { AssetScanner.Scan("relative/path"); throw new Exception("relative path accepted"); }
    catch (ArgumentException) { }
}
finally { Directory.Delete(root, true); }

Console.WriteLine("Asset scanner tests passed");
