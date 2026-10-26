using System.Security.Cryptography;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

string root = Path.Combine(Path.GetTempPath(), "ymm4-asset-scan-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    // A library with >5000 unrelated files must still find media. The old
    // total-entry cap could silently return an empty result in this case.
    for (int i = 0; i < 5200; i++)
        File.WriteAllBytes(Path.Combine(root, $"unrelated-{i:D5}.txt"), Array.Empty<byte>());
    string media = Path.Combine(root, "needle.png");
    byte[] bytes = { 1, 2, 3, 4, 5 };
    File.WriteAllBytes(media, bytes);
    var found = AssetScanner.Scan(root, query: "needle", hash: true);
    Check(found.Success && found.Assets.Length == 1 && found.Assets[0].Name == "needle.png",
        "matching media must be found after unrelated files");
    Check(found.Scanned == 5201 && !found.Truncated,
        "Scanned must report actual inspected entries, without a premature 5000-file cutoff");
    Check(found.Assets[0].Sha256 == Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
        "stable file must receive a correct SHA-256 hash");

    var info = new FileInfo(media);
    DateTime originalWrite = info.LastWriteTimeUtc;
    Check(AssetScanner.HashFile(info, bytes.Length, originalWrite) != null,
        "fresh size and timestamp should certify a digest");
    File.SetLastWriteTimeUtc(media, originalWrite.AddMinutes(-3));
    Check(AssetScanner.HashFile(info, bytes.Length, originalWrite) == null,
        "stale same-length metadata must not certify a digest");
    info.Refresh();
    Check(AssetScanner.HashFile(info, bytes.Length + 1, info.LastWriteTimeUtc) == null,
        "stale file length must not certify a digest");

    File.WriteAllBytes(Path.Combine(root, "duplicate.png"), bytes);
    var duplicates = AssetScanner.Scan(root, hash: true);
    Check(duplicates.Duplicates.Length == 1 && duplicates.Duplicates[0].Length == 2,
        "only matching verified hashes should form duplicate groups");
    var limited = AssetScanner.Scan(root, maxResults: 1);
    Check(limited.Assets.Length == 1 && limited.Truncated,
        "result cap must continue to mark incomplete asset results");
}
finally { Directory.Delete(root, recursive: true); }

Console.WriteLine("Asset scanner integrity tests passed");
