using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace YMM4McpPlugin
{
    /// <summary>
    /// Fail-closed parsing for authenticated native HTTP API request bodies.
    /// The transport applies its byte and time limits before calling Parse.
    /// </summary>
    internal static class RequestBodyValidation
    {
        private static readonly UTF8Encoding StrictUtf8 = new(false, true);

        public static void ValidateLength(long declaredBytes, long receivedBytes)
        {
            // HttpListener reports -1 for chunked requests without Content-Length.
            if (declaredBytes >= 0 && receivedBytes != declaredBytes)
                throw new ArgumentException("Request body length differs from Content-Length");
        }

        public static Dictionary<string, JsonElement> Parse(ReadOnlyMemory<byte> bytes)
        {
            // Preserve the historical behavior of an empty POST body.
            if (bytes.IsEmpty) return new Dictionary<string, JsonElement>(StringComparer.Ordinal);

            // A malformed UTF-8 string must not be silently replaced before a
            // request is interpreted as a destructive edit command.
            try { _ = StrictUtf8.GetCharCount(bytes.Span); }
            catch (DecoderFallbackException ex)
            {
                throw new ArgumentException("Request JSON must be valid UTF-8", ex);
            }

            using var document = JsonDocument.Parse(bytes, new JsonDocumentOptions
            {
                MaxDepth = 32,
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow
            });
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new ArgumentException("Request JSON body must be an object");

            CheckUniqueKeys(document.RootElement);
            var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            foreach (var property in document.RootElement.EnumerateObject())
                result.Add(property.Name, property.Value.Clone());
            return result;
        }

        private static void CheckUniqueKeys(JsonElement element)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                {
                    if (!names.Add(property.Name))
                        throw new ArgumentException("Duplicate JSON property: " + property.Name);
                    CheckUniqueKeys(property.Value);
                }
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    CheckUniqueKeys(item);
            }
        }
    }
}
