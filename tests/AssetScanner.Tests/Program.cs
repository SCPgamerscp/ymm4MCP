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
    // The result order must not depend on the order in which files were created.
    File.WriteAllText(Path.Combine(root, "zeta.png"), "z");
    File.WriteAllText(Path.Combine(root, "alpha.png"), "a");
    var ordered = AssetScanner.Scan(root, recursive: true);
    var names = ordered.Assets.Select(a => a.Path).ToArray();
    Check(names.SequenceEqual(names.OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                   .ThenBy(p => p, StringComparer.Ordinal)),
        "asset results must be deterministic by path");
    var first = AssetScanner.Scan(root, recursive: true, maxResults: 1);
    Check(first.Assets.Length == 1 && first.Assets[0].Name == "alpha.png" && first.Truncated,
        "max_results should select the first sorted matching asset");

    // Locked files must not cause the entire scan to fail. Their hash must not
    // be presented as a verified duplicate when it cannot be read.
    string lockedPath = Path.Combine(root, "locked.wav");
    File.WriteAllText(lockedPath, "locked");
    using (var exclusive = new FileStream(lockedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
    {
        var locked = AssetScanner.Scan(root, "locked", hash: true);
        Check(locked.Success, "locked file must not abort scan");
        Check(locked.Assets.Length == 0 ||
              (locked.Assets.Length == 1 && locked.Assets[0].Sha256 == null && locked.HashTruncated),
              "unreadable files must never receive a misleading SHA256");
    }

    try { AssetScanner.Scan("relative/path"); throw new Exception("relative path accepted"); }
    catch (ArgumentException) { }
}
finally { Directory.Delete(root, true); }

Console.WriteLine("Asset scanner tests passed");
