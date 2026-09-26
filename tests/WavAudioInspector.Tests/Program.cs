using System.Buffers.Binary;
using System.Text;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

static byte[] Wave(short[] samples, int channels = 2, int rate = 8000)
{
    byte[] data = new byte[samples.Length * 2];
    for (int i = 0; i < samples.Length; i++)
        BinaryPrimitives.WriteInt16LittleEndian(data.AsSpan(i * 2, 2), samples[i]);
    var bytes = new byte[44 + data.Length];
    Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), (uint)(bytes.Length - 8));
    Encoding.ASCII.GetBytes("WAVEfmt ").CopyTo(bytes, 8);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(16, 4), 16);
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(20, 2), 1);
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(22, 2), (ushort)channels);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(24, 4), (uint)rate);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(28, 4), (uint)(rate * channels * 2));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(32, 2), (ushort)(channels * 2));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(34, 2), 16);
    Encoding.ASCII.GetBytes("data").CopyTo(bytes, 36);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(40, 4), (uint)data.Length);
    data.CopyTo(bytes, 44);
    return bytes;
}

string path = Path.GetTempFileName();
try
{
    var samples = new short[8000 * 3 * 2];
    for (int frame = 8000 * 2; frame < 8000 * 3; frame++) samples[frame * 2] = 32767;
    File.WriteAllBytes(path, Wave(samples));
    var qa = WavAudioInspector.Inspect(path, 1.5);
    Check(qa.Success && !qa.Passed && qa.DurationSeconds == 3, "valid PCM detection failed");
    Check(qa.Issues.Any(i => i.Code == "LONG_SILENCE" && i.EndSeconds == 2), "silence missing");
    Check(qa.Issues.Any(i => i.Code == "AUDIO_CLIPPING"), "clipping missing");
    Check(qa.Issues.Any(i => i.Code == "CHANNEL_IMBALANCE"), "channel imbalance missing");

    var clean = new short[8000 * 2];
    for (int frame = 0; frame < 8000; frame++)
    {
        clean[frame * 2] = 2000;
        clean[frame * 2 + 1] = 2000;
    }
    File.WriteAllBytes(path, Wave(clean));
    Check(WavAudioInspector.Inspect(path).Passed, "clean stereo should pass");
    File.WriteAllText(path, "invalid wave");
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID", "truncated WAV accepted");
    var unsupported = Wave(clean);
    BinaryPrimitives.WriteUInt16LittleEndian(unsupported.AsSpan(20, 2), 3);
    File.WriteAllBytes(path, unsupported);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FORMAT_UNSUPPORTED", "non-PCM accepted");
}
finally { File.Delete(path); }

Console.WriteLine("WAV audio QA tests passed");
