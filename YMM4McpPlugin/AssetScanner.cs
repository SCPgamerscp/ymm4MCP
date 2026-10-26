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
        // Bound the work even when the requested folder contains no media.
        // Unlike ScanLimit, this counts unrelated files too.
        private const int EnumerationLimit = 20000;
        private const long HashLimit = 64L * 1024 * 1024;
        private const long TotalHashLimit = 256L * 1024 * 1024;
        private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".mp4", ".mov", ".mkv", ".webm", ".avi", ".wav", ".mp3", ".flac", ".ogg",
            ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"
        };

        // The file's metadata must still describe the bytes that were hashed.
        // FileInfo.OpenRead uses FileShare.Read, preventing concurrent writers
        // on Windows while the stream is open. Recheck size, timestamp, and
        // reparse-point status after hashing to reject stale scan metadata.
        internal static string? HashFile(FileInfo file, long expectedBytes, DateTime expectedLastWriteUtc)
        {
            using var stream = file.OpenRead();
            if (stream.Length != expectedBytes) return null;
            string digest = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
            file.Refresh();
            if (!file.Exists || file.Length != expectedBytes
                || file.LastWriteTimeUtc != expectedLastWriteUtc
                || (file.Attributes & FileAttributes.ReparsePoint) != 0)
                return null;
            return digest;
        }

        // Reserve the full candidate size before attempting I/O: a digest may
        // be rejected after hashing (stale metadata), but that work still
        // counts against the scan-wide hashing budget.
        internal static bool TryReserveHashBudget(long bytes, ref long reservedBytes)
        {
            if (bytes < 0 || reservedBytes < 0 || reservedBytes > TotalHashLimit
                || bytes > HashLimit || bytes > TotalHashLimit - reservedBytes)
                return false;
            reservedBytes += bytes;
            return true;
        }

        public static AssetScanResult Scan(string directory, string? query = null, bool recursive = false,
            bool hash = false, int maxResults = 100, int traversalLimit = EnumerationLimit)
        {
            if (string.IsNullOrWhiteSpace(directory) || !Path.IsPathFullyQualified(directory))
                throw new ArgumentException("directory must be an absolute path");
            if (maxResults is < 1 or > 500) throw new ArgumentException("max_results must be in 1..500");
            if (traversalLimit is < 1 or > EnumerationLimit)
                throw new ArgumentOutOfRangeException(nameof(traversalLimit));
            directory = Path.GetFullPath(directory);
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);
            // Enumerate each directory explicitly. Recursive EnumerateFiles only yields
            // files, so a tree of empty directories could previously evade the work cap.
            // An inaccessible subtree must not be silently treated as a complete scan.
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = false,
                IgnoreInaccessible = false,
                AttributesToSkip = FileAttributes.ReparsePoint | FileAttributes.System
            };
            var paths = new List<string>();
            var pendingDirectories = new Queue<string>();
            pendingDirectories.Enqueue(directory);
            int scanned = 0;
            int visitedEntries = 0; // Includes directories, not just files.
            bool truncated = false;
            bool hashTruncated = false;
            bool stop = false;
            while (pendingDirectories.Count > 0 && !stop)
            {
                string current = pendingDirectories.Dequeue();
                try
                {
                    foreach (var entry in Directory.EnumerateFileSystemEntries(current, "*", options))
                    {
                        if (visitedEntries >= traversalLimit)
                        {
                            truncated = true;
                            stop = true;
                            break;
                        }
                        visitedEntries++;
                        FileAttributes attributes;
                        try { attributes = File.GetAttributes(entry); }
                        catch (IOException) { truncated = true; continue; }
                        catch (UnauthorizedAccessException) { truncated = true; continue; }
                        // Recheck attributes: an entry can become a symlink after
                        // enumeration, and must never be followed into another tree.
                        if ((attributes & (FileAttributes.ReparsePoint | FileAttributes.System)) != 0)
                            continue;
                        if ((attributes & FileAttributes.Directory) != 0)
                        {
                            if (recursive) pendingDirectories.Enqueue(entry);
                            continue;
                        }
                        scanned++;
                        if (!Extensions.Contains(Path.GetExtension(entry)) ||
                            (query != null && !Path.GetFileName(entry).Contains(query, StringComparison.OrdinalIgnoreCase)))
                            continue;
                        if (paths.Count >= ScanLimit)
                        {
                            truncated = true;
                            stop = true;
                            break;
                        }
                        paths.Add(entry);
                    }
                }
                catch (IOException) { truncated = true; }
                catch (UnauthorizedAccessException) { truncated = true; }
            }

            var assets = new List<AssetEntry>();
            long hashedBytes = 0;
            foreach (var path in paths.OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                      .ThenBy(p => p, StringComparer.Ordinal))
            {
                if (assets.Count >= maxResults) { truncated = true; break; }
                try
                {
                    // A file can disappear or be replaced by a symlink after
                    // enumeration. Never follow a newly introduced reparse point.
                    var file = new FileInfo(path);
                    if ((file.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                    long bytes = file.Length;
                    DateTime modifiedUtc = file.LastWriteTimeUtc;
                    string modified = modifiedUtc.ToString("O");
                    string? sha = null;
                    if (hash)
                    {
                        if (!TryReserveHashBudget(bytes, ref hashedBytes))
                            hashTruncated = true;
                        else
                        {
                            try
                            {
                                sha = HashFile(file, bytes, modifiedUtc);
                                if (sha == null) hashTruncated = true;
                            }
                            catch (IOException) { hashTruncated = true; }
                            catch (UnauthorizedAccessException) { hashTruncated = true; }
                        }
                    }
                    assets.Add(new AssetEntry(file.FullName, file.Name, file.Extension.ToLowerInvariant(),
                        bytes, modified, sha));
                }
                catch (IOException) { /* File removed while scanning. */ }
                catch (UnauthorizedAccessException) { /* File became inaccessible. */ }
            }
            var duplicates = assets.Where(a => a.Sha256 != null).GroupBy(a => a.Sha256!)
                .Where(g => g.Count() > 1).Select(g => g.Select(a => a.Path).ToArray()).ToArray();
            return new(true, directory, assets.ToArray(), duplicates, scanned, truncated, hashTruncated);
        }
    }
}
