using System.Buffers.Binary;
using System.Text;
using YMM4McpPlugin;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var chunks = new OrderedAudioChunks();
var reusedBuffer = new byte[] { 1, 2, 9 };
chunks.Add(reusedBuffer, 2);
reusedBuffer[0] = 8;
chunks.Add(new byte[] { 3, 4 }, 2);
Check(chunks.ToArray().SequenceEqual(new byte[] { 1, 2, 3, 4 }),
    "WASAPI chunks must preserve callback order and copy reusable buffers");

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

static byte[] PcmWave(int[] samples, int bits, bool extensible = false, int rate = 8000, int? validBits = null)
{
    int sampleBytes = bits / 8, fmtSize = extensible ? 40 : 16, offset = 20 + fmtSize;
    int dataSize = samples.Length * sampleBytes;
    var bytes = new byte[offset + 8 + dataSize + (dataSize & 1)];
    Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), (uint)(bytes.Length - 8));
    Encoding.ASCII.GetBytes("WAVEfmt ").CopyTo(bytes, 8);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(16, 4), (uint)fmtSize);
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(20, 2), (ushort)(extensible ? 0xfffe : 1));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(22, 2), 1);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(24, 4), (uint)rate);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(28, 4), (uint)(rate * sampleBytes));
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(32, 2), (ushort)sampleBytes);
    BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(34, 2), (ushort)bits);
    if (extensible)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(36, 2), 22);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(38, 2), (ushort)(validBits ?? bits));
        new Guid("00000001-0000-0010-8000-00aa00389b71").TryWriteBytes(bytes.AsSpan(44, 16));
    }
    Encoding.ASCII.GetBytes("data").CopyTo(bytes, offset);
    BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 4, 4), (uint)dataSize);
    for (int i = 0; i < samples.Length; i++)
        for (int b = 0; b < sampleBytes; b++)
            bytes[offset + 8 + i * sampleBytes + b] = bits == 8
                ? (byte)(samples[i] + 128)
                : (byte)(samples[i] >> (b * 8));
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
    Check(qa.Issues.Any(i => i.Code == "AUDIO_CLIPPING" && i.StartSeconds == 2 && i.EndSeconds == 3),
        "clipping interval reaching EOF is missing");
    Check(qa.Issues.Any(i => i.Code == "CHANNEL_IMBALANCE"), "channel imbalance missing");
    Check(qa.ChannelLevels.Length == 2
        && qa.ChannelLevels[0].Channel == 1 && qa.ChannelLevels[0].Peak == 32767.0 / 32768
        && Math.Abs(qa.ChannelLevels[0].Rms - (32767.0 / 32768) / Math.Sqrt(3)) < 1e-9
        && qa.ChannelLevels[1] == new AudioQaChannelLevel(2, 0, 0),
        "stereo peak and RMS measurements must include the silent channel");

    var clean = new short[8000 * 2];
    for (int frame = 0; frame < 8000; frame++)
    {
        clean[frame * 2] = 2000;
        clean[frame * 2 + 1] = 2000;
    }
    File.WriteAllBytes(path, Wave(clean));
    Check(WavAudioInspector.Inspect(path).Passed, "clean stereo should pass");

    // The longest run spans two read buffers and is below the whole-file
    // clipped-sample threshold. The right channel is unaffected.
    const int clipStart = 2035, clipLength = 40, testRate = 8000;
    var mostlyClean = new short[testRate * 10 * 2];
    for (int frame = 0; frame < mostlyClean.Length / 2; frame++)
    {
        short value = (short)(frame % 2 == 0 ? 8192 : -8192);
        mostlyClean[frame * 2] = value;
        mostlyClean[frame * 2 + 1] = value;
    }
    var localClip = (short[])mostlyClean.Clone();
    for (int frame = clipStart; frame < clipStart + clipLength; frame++)
        localClip[frame * 2] = short.MaxValue;
    File.WriteAllBytes(path, Wave(localClip));
    var localQa = WavAudioInspector.Inspect(path);
    Check(localQa.Success && !localQa.Passed && localQa.Issues.Any(i =>
        i.Code == "AUDIO_CLIPPING" && i.StartSeconds == (double)clipStart / testRate
        && i.EndSeconds == (double)(clipStart + clipLength) / testRate),
        "local clipping must report its exact interval across buffers");

    // Alternating which channel reaches full scale is not a continuous
    // clipping run on either channel.
    for (int frame = clipStart; frame < clipStart + clipLength; frame++)
        mostlyClean[frame * 2 + frame % 2] = short.MaxValue;
    File.WriteAllBytes(path, Wave(mostlyClean));
    Check(WavAudioInspector.Inspect(path).Passed,
        "alternating channels must not count as a continuous clipping run");
    File.WriteAllBytes(path, Wave(new short[8000 * 2]));
    var shortSilence = WavAudioInspector.Inspect(path, 2);
    Check(shortSilence.Success && !shortSilence.Passed &&
        shortSilence.Issues.Any(i => i.Code == "AUDIO_SILENT"),
        "a fully silent clip shorter than min_silence_seconds must fail");
    var inverted = new short[8000 * 2];
    for (int frame = 0; frame < 8000; frame++)
    {
        inverted[frame * 2] = (short)(frame % 2 == 0 ? 8000 : -8000);
        inverted[frame * 2 + 1] = (short)-inverted[frame * 2];
    }
    File.WriteAllBytes(path, Wave(inverted));
    var phaseQa = WavAudioInspector.Inspect(path);
    Check(phaseQa.Passed && phaseQa.Issues.Any(i => i.Code == "CHANNEL_PHASE_INVERSION"),
        "opposite-phase stereo should be flagged as a warning");
    Array.Fill(inverted, (short)12000);
    File.WriteAllBytes(path, Wave(inverted));
    Check(WavAudioInspector.Inspect(path).Issues.Any(i => i.Code == "DC_OFFSET"),
        "large DC offset should be flagged");
    File.WriteAllBytes(path, Wave(clean)[..20]);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID", "truncated WAV accepted");
    var unsupported = Wave(clean);
    BinaryPrimitives.WriteUInt16LittleEndian(unsupported.AsSpan(20, 2), 3);
    File.WriteAllBytes(path, unsupported);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FORMAT_UNSUPPORTED", "mismatched float format accepted");

    // RIFF fmt metadata must be self-consistent; otherwise audio QA could accept corrupt input.
    var invalidByteRate = Wave(clean);
    BinaryPrimitives.WriteUInt32LittleEndian(invalidByteRate.AsSpan(28, 4), 1);
    File.WriteAllBytes(path, invalidByteRate);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "inconsistent WAV byte rate accepted");
    var invalidSampleRate = Wave(clean);
    BinaryPrimitives.WriteUInt32LittleEndian(invalidSampleRate.AsSpan(24, 4), 16000);
    File.WriteAllBytes(path, invalidSampleRate);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "sample rate changed without matching byte rate accepted");
    var originalWave = Wave(clean);
    var fmtChunk = originalWave.AsSpan(12, 24).ToArray();
    var duplicateFmt = new byte[originalWave.Length + fmtChunk.Length];
    originalWave.CopyTo(duplicateFmt, 0);
    fmtChunk.CopyTo(duplicateFmt, originalWave.Length);
    BinaryPrimitives.WriteUInt32LittleEndian(duplicateFmt.AsSpan(4, 4), (uint)(duplicateFmt.Length - 8));
    File.WriteAllBytes(path, duplicateFmt);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "duplicate fmt chunk accepted");

    // A 17-byte fmt chunk has an incomplete cbSize field. An 18-byte
    // WAVEFORMATEX header may not claim nonexistent extension bytes.
    var basePcm = Wave(clean);
    var extendedPcm = new byte[basePcm.Length + 2];
    Array.Copy(basePcm, 0, extendedPcm, 0, 36);
    Array.Copy(basePcm, 36, extendedPcm, 38, basePcm.Length - 36);
    BinaryPrimitives.WriteUInt32LittleEndian(extendedPcm.AsSpan(4, 4),
        (uint)(extendedPcm.Length - 8));
    BinaryPrimitives.WriteUInt32LittleEndian(extendedPcm.AsSpan(16, 4), 17);
    File.WriteAllBytes(path, extendedPcm);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "incomplete WAVEFORMATEX cbSize accepted");
    BinaryPrimitives.WriteUInt32LittleEndian(extendedPcm.AsSpan(16, 4), 18);
    BinaryPrimitives.WriteUInt16LittleEndian(extendedPcm.AsSpan(36, 2), 1);
    File.WriteAllBytes(path, extendedPcm);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "WAVEFORMATEX extension length beyond fmt accepted");
    BinaryPrimitives.WriteUInt16LittleEndian(extendedPcm.AsSpan(36, 2), 0);
    File.WriteAllBytes(path, extendedPcm);
    Check(WavAudioInspector.Inspect(path).Passed,
        "valid WAVEFORMATEX cbSize=0 must remain supported");

    // A second data chunk cannot be silently excluded from the QA duration.
    var duplicateData = new byte[basePcm.Length + 10];
    basePcm.CopyTo(duplicateData, 0);
    Encoding.ASCII.GetBytes("data").CopyTo(duplicateData, basePcm.Length);
    BinaryPrimitives.WriteUInt32LittleEndian(duplicateData.AsSpan(basePcm.Length + 4, 4), 2);
    BinaryPrimitives.WriteInt16LittleEndian(duplicateData.AsSpan(basePcm.Length + 8, 2), 1000);
    BinaryPrimitives.WriteUInt32LittleEndian(duplicateData.AsSpan(4, 4),
        (uint)(duplicateData.Length - 8));
    File.WriteAllBytes(path, duplicateData);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "duplicate data chunk was silently ignored");


    // Extensible fmt cbSize must describe bytes actually present in the fmt chunk.
    var invalidExtensibleLength = PcmWave(new[] { 1, 2, 3, 4 }, 24, extensible: true);
    BinaryPrimitives.WriteUInt16LittleEndian(invalidExtensibleLength.AsSpan(36, 2), 23);
    File.WriteAllBytes(path, invalidExtensibleLength);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "extensible fmt claiming missing extension bytes accepted");

    // RIFF payload must not silently ignore 1..7 trailing bytes of a chunk header.
    var partialHeader = Wave(clean);
    Array.Resize(ref partialHeader, partialHeader.Length + 3);
    BinaryPrimitives.WriteUInt32LittleEndian(partialHeader.AsSpan(4, 4), (uint)(partialHeader.Length - 8));
    File.WriteAllBytes(path, partialHeader);
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID",
        "truncated trailing RIFF chunk header accepted");

    foreach (int bits in new[] { 8, 24, 32 })
        foreach (bool extensible in new[] { false, true })
        {
            int halfScale = 1 << (bits - 2);
            var pcmSamples = Enumerable.Range(0, 8000).Select(i => i % 2 == 0 ? halfScale : -halfScale).ToArray();
            File.WriteAllBytes(path, PcmWave(pcmSamples, bits, extensible));
            var pcmQa = WavAudioInspector.Inspect(path);
            Check(pcmQa.Passed && pcmQa.Peak == 0.5 && pcmQa.Rms == 0.5,
                $"PCM {bits}-bit extensible={extensible} QA failed");
            Check(pcmQa.ChannelLevels.SequenceEqual(new[] { new AudioQaChannelLevel(1, 0.5, 0.5) }),
                $"PCM {bits}-bit mono channel levels incorrect");
        }
    foreach (int rail in new[] { -128, 127 })
    {
        File.WriteAllBytes(path, PcmWave(new[] { rail, 0, 0 }, 8));
        var pcm8Qa = WavAudioInspector.Inspect(path);
        Check(pcm8Qa.Success && pcm8Qa.DurationSeconds == 3.0 / 8000
            && pcm8Qa.Issues.Any(i => i.Code == "AUDIO_CLIPPING"),
            "PCM8 must detect both clipping endpoints without including RIFF padding as audio");
    }
    foreach (var (bits, validBits, positiveRail, nextNegative) in new[]
        { (8, 8, 127, -127), (8, 7, 126, -126), (16, 8, 32512, -32512) })
    {
        File.WriteAllBytes(path, PcmWave(Enumerable.Repeat(positiveRail, 8000).ToArray(),
            bits, extensible: true, validBits: validBits));
        var clippedQa = WavAudioInspector.Inspect(path);
        Check(clippedQa.Success && !clippedQa.Passed && clippedQa.Issues.Any(i =>
            i.Code == "AUDIO_CLIPPING" && i.StartSeconds == 0 && i.EndSeconds == 1),
            $"PCM {validBits}/{bits} positive endpoint must be detected");
        File.WriteAllBytes(path, PcmWave(Enumerable.Repeat(nextNegative, 8000).ToArray(),
            bits, extensible: true, validBits: validBits));
        Check(WavAudioInspector.Inspect(path).Passed,
            $"PCM {validBits}/{bits} next negative code must keep the common negative threshold");
    }
    File.WriteAllBytes(path, PcmWave(new int[80], 8, extensible: true, validBits: 1));
    var oneBitSilence = WavAudioInspector.Inspect(path);
    Check(oneBitSilence.Success && oneBitSilence.Issues.Any(i => i.Code == "AUDIO_SILENT")
        && !oneBitSilence.Issues.Any(i => i.Code == "AUDIO_CLIPPING"),
        "one-bit silence must not be counted as clipping");
    var floatSamples = new float[8000 * 2];
    Array.Fill(floatSamples, 0.25f);
    foreach (bool extensible in new[] { false, true })
    {
        File.WriteAllBytes(path, FloatWave(floatSamples, extensible));
        var floatQa = WavAudioInspector.Inspect(path);
        Check(floatQa.Passed && floatQa.Peak == 0.25 && floatQa.Rms == 0.25,
            "IEEE float mono/stereo QA failed");

        const int firstPeakFrame = 4097;
        int peakIndex = firstPeakFrame * 2 + 1;
        floatSamples[peakIndex] = extensible ? -1f : 1f;
        File.WriteAllBytes(path, FloatWave(floatSamples, extensible));
        Check(WavAudioInspector.Inspect(path).Passed,
            "a single sample exactly at full scale must keep the existing clipping threshold");

        floatSamples[peakIndex] = extensible ? -1.25f : MathF.BitIncrement(1f);
        File.WriteAllBytes(path, FloatWave(floatSamples, extensible));
        var overRangeQa = WavAudioInspector.Inspect(path);
        Check(overRangeQa.Success && !overRangeQa.Passed && overRangeQa.Issues.Count(i =>
            i.Code == "AUDIO_CLIPPING") == 1 && overRangeQa.Issues.Any(i =>
            i.Code == "AUDIO_CLIPPING" && i.StartSeconds == firstPeakFrame / 8000.0 &&
            i.EndSeconds == (firstPeakFrame + 1) / 8000.0),
            "an isolated float sample beyond full scale must fail QA at its actual frame");
        floatSamples[peakIndex] = 0.25f;
    }
    floatSamples[0] = float.NaN;
    File.WriteAllBytes(path, FloatWave(floatSamples));
    Check(WavAudioInspector.Inspect(path).ErrorCode == "AUDIO_FILE_INVALID", "NaN float accepted");
}
finally { File.Delete(path); }

Console.WriteLine("WAV audio QA tests passed");
