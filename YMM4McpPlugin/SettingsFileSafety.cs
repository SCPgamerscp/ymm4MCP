using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace YMM4McpPlugin
{
    // Settings control the listener and advanced APIs. Malformed or ambiguous
    // settings must never be silently accepted as a successful configuration.
    internal static class SettingsFileSafety
    {
        internal const int MaxSettingsBytes = 64 * 1024;
        private static readonly object SaveLock = new();

        internal static string ReadValidated(string path)
        {
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            // Read no more than one byte beyond the limit, even if the file grows
            // while it is being read. Avoid allocating based on untrusted file size.
            var buffer = new byte[MaxSettingsBytes + 1];
            int length = 0;
            while (length < buffer.Length)
            {
                int count = file.Read(buffer, length, buffer.Length - length);
                if (count == 0) break;
                length += count;
            }
            if (length > MaxSettingsBytes)
                throw new InvalidDataException("MCP設定ファイルが64KiBを超えています");

            // Do not replace malformed UTF-8 bytes with U+FFFD in a security-
            // sensitive configuration. Preserve the BOM supported by ReadAllText.
            string json;
            try { json = new UTF8Encoding(false, true).GetString(buffer, 0, length); }
            catch (DecoderFallbackException ex)
            {
                // McpSettings.Load catches InvalidDataException and falls back to
                // AutoStart=false with a visible warning instead of breaking UI init.
                throw new InvalidDataException("MCP設定ファイルが正しいUTF-8ではありません", ex);
            }
            if (json.Length > 0 && json[0] == '\uFEFF')
                json = json.Substring(1);

            using var document = JsonDocument.Parse(json);
            RejectDuplicateKeys(document.RootElement);
            // JsonSerializer ignores unknown properties by default. In settings
            // controlling the listener, a typo must not silently fall back to
            // AutoStart=true or a different port.
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException("MCP設定JSONはオブジェクトで指定してください");
            foreach (var property in document.RootElement.EnumerateObject())
                if (property.Name is not ("Port" or "AutoStart" or "AllowAdvanced"))
                    throw new InvalidDataException("MCP設定JSONに不明なキーがあります: " + property.Name);
            return json;
        }

        private static void RejectDuplicateKeys(JsonElement element)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                {
                    if (!names.Add(property.Name))
                        throw new InvalidDataException("MCP設定JSONに重複したキーがあります: " + property.Name);
                    RejectDuplicateKeys(property.Value);
                }
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    RejectDuplicateKeys(item);
            }
        }

        internal static void WriteAtomic(string path, string json)
        {
            // Unique same-directory staging files prevent simultaneous saves from
            // clobbering each other's .tmp file. Never expose partial JSON.
            string staging = path + ".tmp." + Guid.NewGuid().ToString("N");
            try
            {
                using (var file = new FileStream(staging, FileMode.CreateNew,
                           FileAccess.Write, FileShare.None))
                {
                    using var writer = new StreamWriter(file, new UTF8Encoding(false),
                        bufferSize: 4096, leaveOpen: true);
                    writer.Write(json);
                    writer.Flush();
                    file.Flush(flushToDisk: true);
                }
                // Windows may deny overlapping replacements of the same
                // destination even when each writer has a unique staging file.
                lock (SaveLock)
                    File.Move(staging, path, overwrite: true);
            }
            finally
            {
                try { File.Delete(staging); }
                catch (IOException) { /* Preserve original failure. */ }
                catch (UnauthorizedAccessException) { /* Preserve original failure. */ }
            }
        }
    }
}
