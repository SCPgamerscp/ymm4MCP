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

static byte[] FloatWave(float[] samples, bool extensible = false, int channels = 2, int rate = 8000)
{
    int fmtSize = extensible ? 40 : 16;
    int offset = 20 + fmtSize;
    byte[] bytes = new byte[offset + 8 + samples.Length * 4];
    Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), (uint)(bytes.Length - 8));
    Encoding.ASCII.GetBytes("WAVEfmt ").CopyTo(bytes, 8);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(16, 4), (uint)fmtSize);
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(20, 2), (ushort)(extensible ? 0xfffe : 3));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(22, 2), (ushort)channels);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(24, 4), (uint)rate);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(28, 4), (uint)(rate * channels * 4));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(32, 2), (ushort)(channels * 4));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(34, 2), 32);
    if (extensible)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(36, 2), 22);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(38, 2), 32);
        new Guid("00000003-0000-0010-8000-00aa00389b71").TryWriteBytes(bytes.AsSpan(44, 16));
    }
    Encoding.ASCII.GetBytes("data").CopyTo(bytes, offset);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 4, 4), (uint)(samples.Length * 4));
    for (int i = 0; i < samples.Length; i++)
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(offset + 8 + i * 4, 4), BitConverter.SingleToInt32Bits(samples[i]));
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
    File.WriteAllBytes(path, Wave(clean)[..20]);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID", "truncated WAV accepted");
    var unsupported = Wave(clean);
    BinaryPrimitives.WriteUInt16LittleEndian(unsupported.AsSpan(20, 2), 3);
    File.WriteAllBytes(path, unsupported);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FORMAT_UNSUPPORTED", "mismatched float format accepted");
    var floatSamples = new float[8000 * 2];
    Array.Fill(floatSamples, 0.25f);
    foreach (bool extensible in new[] { false, true })
    {
        File.WriteAllBytes(path, FloatWave(floatSamples, extensible));
        var floatQa = WavAudioInspector.Inspect(path);
        Check(floatQa.Passed && floatQa.Peak == 0.25 && floatQa.Rms == 0.25,
            "IEEE float mono/stereo QA failed");
    }
    floatSamples[0] = float.NaN;
    File.WriteAllBytes(path, FloatWave(floatSamples));
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID", "NaN float accepted");
}
finally { File.Delete(path); }

Console.WriteLine("WAV audio QA tests passed");
