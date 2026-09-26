using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace YMM4McpPlugin
{
    internal sealed record AssetEntry(string Path, string Name, string Extension, long Bytes,
        string LastWriteUtc, string? Sha256);
    internal sealed record AssetScanResult(bool Success, string Directory, AssetEntry[] Assets,
        string[][] Duplicates, int Scanned, bool Truncated, bool HashTruncated);

    internal static class AssetScanner
    {
        private const int ScanLimit = 5000;
        private const long HashLimit = 64L * 1024 * 1024;
        private const long TotalHashLimit = 256L * 1024 * 1024;
        private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".mp4", ".mov", ".mkv", ".webm", ".avi", ".wav", ".mp3", ".flac", ".ogg",
            ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"
        };

        public static AssetScanResult Scan(string directory, string? query = null, bool recursive = false,
            bool hash = false, int maxResults = 100)
        {
            if (string.IsNullOrWhiteSpace(directory) || !Path.IsPathFullyQualified(directory))
                throw new ArgumentException("directory must be an absolute path");
            if (maxResults is < 1 or > 500) throw new ArgumentException("max_results must be in 1..500");
            directory = Path.GetFullPath(directory);
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = recursive,
                IgnoreInaccessible = true,
                AttributesToSkip = FileAttributes.ReparsePoint | FileAttributes.System
            };
            var assets = new List<AssetEntry>();
            int scanned = 0;
            bool truncated = false;
            bool hashTruncated = false;
            long hashedBytes = 0;
            foreach (var path in Directory.EnumerateFiles(directory, "*", options))
            {
                if (++scanned > ScanLimit) { truncated = true; break; }
                var file = new FileInfo(path);
                if (!Extensions.Contains(file.Extension) ||
                    (query != null && !file.Name.Contains(query, StringComparison.OrdinalIgnoreCase))) continue;
                if (assets.Count >= maxResults) { truncated = true; break; }
                string? sha = null;
                if (hash && file.Length <= HashLimit && file.Length <= TotalHashLimit - hashedBytes)
                {
                    using var stream = file.OpenRead();
                    sha = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
                    hashedBytes += file.Length;
                }
                else if (hash) hashTruncated = true;
                assets.Add(new AssetEntry(file.FullName, file.Name, file.Extension.ToLowerInvariant(),
                    file.Length, file.LastWriteTimeUtc.ToString("O"), sha));
            }
            var duplicates = assets.Where(a => a.Sha256 != null).GroupBy(a => a.Sha256!)
                .Where(g => g.Count() > 1).Select(g => g.Select(a => a.Path).ToArray()).ToArray();
            return new(true, directory, assets.ToArray(), duplicates, scanned, truncated, hashTruncated);
        }
    }
}
