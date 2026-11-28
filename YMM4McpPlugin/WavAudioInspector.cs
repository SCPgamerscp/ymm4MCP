using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace YMM4McpPlugin
{
    internal sealed record AudioQaIssue(string Code, string Severity, double? StartSeconds,
        double? EndSeconds, string Message);
    internal sealed record AudioQaChannelLevel(int Channel, double Peak, double Rms);
    internal sealed record AudioQaReport(bool Success, bool Passed, string? ErrorCode, string? Error,
        int? SampleRate, int? Channels, double? DurationSeconds, double? Peak,
        double? Rms, AudioQaIssue[] Issues, AudioQaChannelLevel[] ChannelLevels);

    internal static class WavAudioInspector
    {
        private const double MinContinuousClipSeconds = 0.005;
        private const double NearFullScale = 32760.0 / 32768;

        public static AudioQaReport Inspect(string path, double minSilenceSeconds = 2)
        {
            if (!double.IsFinite(minSilenceSeconds) || minSilenceSeconds is < 0.1 or > 60)
                throw new ArgumentException("min_silence_seconds must be in 0.1..60");
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
                throw new ArgumentException("path must be an absolute file path");
            if (!File.Exists(path)) return Fail("FILE_NOT_FOUND", "音声ファイルがありません");
            try
            {
                using var file = File.OpenRead(path);
                var riff = Read(file, 12);
                if (Text(riff, 0) != "RIFF" || Text(riff, 8) != "WAVE")
                    return Fail("AUDIO_FORMAT_UNSUPPORTED", "RIFF/WAVE 形式が必要です");
                long end = 8L + U32(riff, 4);
                if (end > file.Length || end < 44) return Fail("AUDIO_FILE_INVALID", "WAV のサイズが不正です");
                int rate = 0, channels = 0, align = 0, bits = 0, validBits = 0, format = 0;
                uint byteRate = 0;
                bool fmtSeen = false;
                long dataStart = -1, dataSize = 0;
                int chunks = 0;
                while (file.Position + 8 <= end)
                {
                    if (++chunks > 1000) return Fail("AUDIO_FILE_INVALID", "WAV のチャンク数が多すぎます");
                    var header = Read(file, 8);
                    long size = U32(header, 4), next = file.Position + size;
                    if (next > end || next + (size & 1) > end)
                        return Fail("AUDIO_FILE_INVALID", "WAV のチャンク長が不正です");
                    if (Text(header, 0) == "fmt ")
                    {
                        if (fmtSeen) return Fail("AUDIO_FILE_INVALID", "fmt チャンクが重複しています");
                        fmtSeen = true;
                        if (size < 16) return Fail("AUDIO_FILE_INVALID", "fmt チャンクが不完全です");
                        var fmt = Read(file, (int)Math.Min(size, 40));
                        // WAVEFORMATEX extensions need both cbSize bytes, and
                        // cbSize cannot claim bytes outside the fmt chunk.
                        if (size > 16 && (size < 18 || U16(fmt, 16) > size - 18))
                            return Fail("AUDIO_FILE_INVALID", "WAV fmt の拡張長が不正です");
                        format = U16(fmt, 0); channels = U16(fmt, 2); rate = checked((int)U32(fmt, 4));
                        byteRate = U32(fmt, 8);
                        align = U16(fmt, 12); bits = U16(fmt, 14);
                        validBits = bits;
                        // WAVE_FORMAT_EXTENSIBLE's cbSize must fit within the fmt chunk.
                        // Otherwise malformed headers can claim extension bytes that do not exist.
                        if (format == 0xfffe &&
                            (size < 40 || U16(fmt, 16) < 22 || U16(fmt, 16) > size - 18))
                            return Fail("AUDIO_FILE_INVALID", "WAVE_FORMAT_EXTENSIBLE fmt の長さが不正です");
                        // WAVE_FORMAT_EXTENSIBLE carries the actual PCM/float format in its subtype GUID.
                        if (format == 0xfffe)
                        {
                            int precision = U16(fmt, 18);
                            if (precision > 0 && precision <= bits)
                            {
                                var subtype = new Guid(fmt.AsSpan(24, 16));
                                if (subtype == new Guid("00000001-0000-0010-8000-00aa00389b71"))
                                {
                                    format = 1;
                                    validBits = precision;
                                }
                                else if (bits == 32 && precision == 32 &&
                                    subtype == new Guid("00000003-0000-0010-8000-00aa00389b71")) format = 3;
                            }
                        }
                    }
                    else if (Text(header, 0) == "data")
                    {
                        // A second data chunk must not be silently ignored:
                        // that would report QA only for the first audio segment.
                        if (dataStart >= 0)
                            return Fail("AUDIO_FILE_INVALID", "WAV data チャンクが重複しています");
                        dataStart = file.Position;
                        dataSize = size;
                    }
                    file.Position = next + (size & 1);
                }
                // A partial chunk header at the end of the RIFF payload is corruption,
                // not a valid end-of-chunks marker.
                if (file.Position != end)
                    return Fail("AUDIO_FILE_INVALID", "WAV のチャンクヘッダーが不完全です");
                bool pcm = format == 1 && bits is 8 or 16 or 24 or 32;
                bool ieeeFloat = format == 3 && bits == 32;
                if ((!pcm && !ieeeFloat) || channels is < 1 or > 2 || rate is < 1 or > 384000 ||
                    align != channels * (bits / 8))
                    return Fail("AUDIO_FORMAT_UNSUPPORTED", "PCM 8/16/24/32-bit / IEEE float 32-bit mono/stereo WAV のみ検査できます");
                // A malformed byte rate can make downstream duration and QA results unreliable.
                if (byteRate != (long)rate * align)
                    return Fail("AUDIO_FILE_INVALID", "WAV の byte rate と sample rate/block align が一致しません");
                if (dataStart < 0 || dataSize == 0 || dataSize % align != 0)
                    return Fail("AUDIO_FILE_INVALID", "PCM データが不完全です");

                long frames = dataSize / align;
                long minSilenceFrames = (long)Math.Ceiling(minSilenceSeconds * rate);
                long minClipFrames = Math.Max(2, (long)Math.Ceiling(MinContinuousClipSeconds * rate));
                // WAVE_FORMAT_EXTENSIBLE stores valid PCM bits at the high end
                // of each sample container. Padding must not affect measurements.
                int pcmMask = -1 << (bits - validBits);
                double pcmScale = 1L << (bits - 1);
                // Left-aligned PCM may reach its positive endpoint below the common
                // threshold. One-bit PCM has no positive code above silence.
                double positiveClipThreshold = pcm && validBits > 1
                    ? Math.Min(NearFullScale, 1.0 - 1.0 / (1L << (validBits - 1)))
                    : NearFullScale;
                var issues = new List<AudioQaIssue>();
                double[] sumSquares = new double[channels];
                double[] sumSamples = new double[channels];
                double[] peaksByChannel = new double[channels];
                long[] clipRuns = new long[channels];
                long[] clippedSamples = new long[channels];
                double crossSamples = 0;
                long silentStart = -1;
                long longestClipLength = 0, longestClipEnd = 0;
                long firstOverRangeFrame = -1;
                double peak = 0;
                file.Position = dataStart;
                var buffer = new byte[8192];
                long current = 0;
                while (current < frames)
                {
                    int count = (int)Math.Min(buffer.Length / align, frames - current);
                    file.ReadExactly(buffer.AsSpan(0, count * align));
                    for (int i = 0; i < count; i++)
                    {
                        long frame = current + i;
                        bool silent = true;
                        double firstChannel = 0;
                        for (int ch = 0; ch < channels; ch++)
                        {
                            double normalized;
                            int offset = i * align + ch * (bits / 8);
                            if (pcm)
                            {
                                int sample = bits switch
                                {
                                    8 => buffer[offset] - 128,
                                    16 => BinaryPrimitives.ReadInt16LittleEndian(buffer.AsSpan(offset, 2)),
                                    24 => (buffer[offset] | buffer[offset + 1] << 8 |
                                        (sbyte)buffer[offset + 2] << 16),
                                    _ => BinaryPrimitives.ReadInt32LittleEndian(buffer.AsSpan(offset, 4))
                                };
                                normalized = (sample & pcmMask) / pcmScale;
                            }
                            else
                            {
                                normalized = BitConverter.Int32BitsToSingle(
                                    BinaryPrimitives.ReadInt32LittleEndian(buffer.AsSpan(offset, 4)));
                                if (!double.IsFinite(normalized))
                                    return Fail("AUDIO_FILE_INVALID", "非有限値の音声サンプルがあります");
                            }
                            double amplitude = Math.Abs(normalized);
                            peak = Math.Max(peak, amplitude);
                            peaksByChannel[ch] = Math.Max(peaksByChannel[ch], amplitude);
                            if (ieeeFloat && amplitude > 1 && firstOverRangeFrame < 0)
                                firstOverRangeFrame = frame;
                            if (normalized <= -NearFullScale || normalized >= positiveClipThreshold)
                            {
                                clippedSamples[ch]++;
                                long runLength = ++clipRuns[ch];
                                if (runLength > longestClipLength)
                                {
                                    longestClipLength = runLength;
                                    longestClipEnd = frame + 1;
                                }
                            }
                            else
                                clipRuns[ch] = 0;
                            if (amplitude > 104.0 / 32768) silent = false; // about -50 dBFS
                            sumSquares[ch] += normalized * normalized;
                            sumSamples[ch] += normalized;
                            if (ch == 0) firstChannel = normalized;
                            else crossSamples += firstChannel * normalized;
                        }
                        if (silent && silentStart < 0) silentStart = frame;
                        if (!silent && silentStart >= 0)
                        {
                            AddSilence(issues, silentStart, frame, minSilenceFrames, rate);
                            silentStart = -1;
                        }
                    }
                    current += count;
                }
                if (silentStart >= 0) AddSilence(issues, silentStart, frames, minSilenceFrames, rate);
                // A completely silent short clip must not pass merely because it
                // is shorter than the configured minimum silence interval.
                if (peak <= 104.0 / 32768 && !issues.Any(i => i.Code == "LONG_SILENCE"))
                    issues.Add(new("AUDIO_SILENT", "error", 0, (double)frames / rate,
                        "音声全体が無音または極めて小さい音量です"));
                if (longestClipLength >= minClipFrames)
                    issues.Add(new("AUDIO_CLIPPING", "error",
                        (double)(longestClipEnd - longestClipLength) / rate, (double)longestClipEnd / rate,
                        "フルスケール付近のサンプルが続いた最長区間です"));
                else if (firstOverRangeFrame >= 0)
                    issues.Add(new("AUDIO_CLIPPING", "error",
                        (double)firstOverRangeFrame / rate, (double)(firstOverRangeFrame + 1) / rate,
                        "浮動小数点 WAV でフルスケールを超えた最初のフレームです"));
                else
                {
                    // Check the 0.1% limit per channel so a clean channel cannot
                    // dilute the clipped-sample fraction of the other channel.
                    var clippedChannels = Enumerable.Range(0, channels)
                        .Where(ch => clippedSamples[ch] > frames / 1000)
                        .Select(ch => ch + 1).ToArray();
                    if (clippedChannels.Length > 0)
                        issues.Add(new("AUDIO_CLIPPING", "error", null, null,
                            $"チャンネル {string.Join("、", clippedChannels)} でフルスケール付近のサンプルが0.1%を超えています"));
                }
                var rmsByChannel = sumSquares.Select(sum => Math.Sqrt(sum / frames)).ToArray();
                var means = sumSamples.Select(sum => sum / frames).ToArray();
                for (int ch = 0; ch < channels; ch++)
                    if (Math.Abs(means[ch]) > 0.1 && Math.Abs(means[ch]) > rmsByChannel[ch] * 0.5)
                        issues.Add(new("DC_OFFSET", "warning", null, null,
                            $"チャンネル {ch + 1} の直流成分が大きい可能性があります"));
                if (channels == 2 && rmsByChannel.Max() > 0.01 &&
                    rmsByChannel.Min() * 4 < rmsByChannel.Max())
                    issues.Add(new("CHANNEL_IMBALANCE", "warning", null, null, "左右チャンネルの RMS に大きな差があります"));
                if (channels == 2)
                {
                    double varianceLeft = sumSquares[0] / frames - means[0] * means[0];
                    double varianceRight = sumSquares[1] / frames - means[1] * means[1];
                    if (varianceLeft > 0.0001 && varianceRight > 0.0001 &&
                        (crossSamples / frames - means[0] * means[1]) /
                        Math.Sqrt(varianceLeft * varianceRight) < -0.95)
                        issues.Add(new("CHANNEL_PHASE_INVERSION", "warning", null, null,
                            "左右チャンネルがほぼ逆位相です"));
                }
                var levels = peaksByChannel.Select((value, channel) =>
                    new AudioQaChannelLevel(channel + 1, value, rmsByChannel[channel])).ToArray();
                return new(true, !issues.Any(i => i.Severity == "error"), null, null, rate, channels,
                    (double)frames / rate, peak,
                    Math.Sqrt(sumSquares.Sum() / (frames * channels)), issues.ToArray(), levels);
            }
            catch (Exception ex) when (ex is IOException or OverflowException or ArgumentException)
            {
                return Fail("AUDIO_FILE_INVALID", "WAV を読み取れません: " + ex.Message);
            }
        }

        private static void AddSilence(List<AudioQaIssue> issues, long start, long end, long minimum, int rate)
        {
            if (end - start >= minimum && issues.Count < 100)
                issues.Add(new("LONG_SILENCE", "error", (double)start / rate, (double)end / rate,
                    "指定秒数以上の無音があります"));
        }

        private static AudioQaReport Fail(string code, string message)
            => new(false, false, code, message, null, null, null, null, null,
                Array.Empty<AudioQaIssue>(), Array.Empty<AudioQaChannelLevel>());
        private static byte[] Read(Stream stream, int count)
        {
            var bytes = new byte[count];
            stream.ReadExactly(bytes);
            return bytes;
        }
        private static string Text(byte[] bytes, int offset) => Encoding.ASCII.GetString(bytes, offset, 4);
        private static ushort U16(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));
        private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
    }
}
