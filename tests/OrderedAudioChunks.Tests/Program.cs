using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var chunks = new OrderedAudioChunks(maxBytes: 4);
var source = new byte[] { 1, 2, 3, 4 };
chunks.Add(source, 2);
source[0] = 9; // WASAPI reuses its callback buffers.
chunks.Add(new byte[] { 3, 4 }, 2);
Check(!chunks.IsTruncated, "Exact limit must be accepted");
Check(chunks.ToArray().SequenceEqual(new byte[] { 1, 2, 3, 4 }),
    "Callback data must be copied in order");
chunks.Add(new byte[] { 5 }, 1);
Check(chunks.IsTruncated, "Exceeding the byte budget must be reported");
chunks.Add(new byte[] { 6 }, 1);
Check(chunks.ToArray().SequenceEqual(new byte[] { 1, 2, 3, 4 }),
    "No additional data may be appended after truncation");

var zero = new OrderedAudioChunks(maxBytes: 1);
zero.Add(Array.Empty<byte>(), 0);
Check(!zero.IsTruncated && zero.ToArray().Length == 0,
    "Empty callbacks must not affect capture state");
bool invalidLength = false;
try { zero.Add(new byte[] { 1 }, 2); }
catch (ArgumentOutOfRangeException) { invalidLength = true; }
Check(invalidLength, "Invalid callback lengths must still be rejected");
bool invalidSource = false;
try { zero.Add(null!, 1); }
catch (ArgumentNullException) { invalidSource = true; }
Check(invalidSource, "Null callback buffers must be rejected explicitly");
Console.WriteLine("Ordered audio chunks regression tests passed");
