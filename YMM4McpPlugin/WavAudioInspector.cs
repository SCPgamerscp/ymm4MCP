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
    internal sealed record AudioQaReport(bool Success, bool Passed, string? ErrorCode, string? Error,
        int? SampleRate, int? Channels, double? DurationSeconds, double? Peak,
        double? Rms, AudioQaIssue[] Issues);

    internal static class WavAudioInspector
    {
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
                int rate = 0, channels = 0, align = 0, bits = 0, format = 0;
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
                        if (size < 16) return Fail("AUDIO_FILE_INVALID", "fmt チャンクが不完全です");
                        var fmt = Read(file, 16);
                        format = U16(fmt, 0); channels = U16(fmt, 2); rate = checked((int)U32(fmt, 4));
                        align = U16(fmt, 12); bits = U16(fmt, 14);
                    }
                    else if (Text(header, 0) == "data" && dataStart < 0)
                    {
                        dataStart = file.Position;
                        dataSize = size;
                    }
                    file.Position = next + (size & 1);
                }
                if (format != 1 || channels is < 1 or > 2 || bits != 16 || rate is < 1 or > 384000 || align != channels * 2)
                    return Fail("AUDIO_FORMAT_UNSUPPORTED", "PCM 16-bit mono/stereo WAV のみ検査できます");
                if (dataStart < 0 || dataSize == 0 || dataSize % align != 0)
                    return Fail("AUDIO_FILE_INVALID", "PCM データが不完全です");

                long frames = dataSize / align;
                long minSilenceFrames = (long)Math.Ceiling(minSilenceSeconds * rate);
                var issues = new List<AudioQaIssue>();
                double[] sumSquares = new double[channels];
                long clipped = 0, silentStart = -1;
                int peak = 0;
                file.Position = dataStart;
                var buffer = new byte[8192];
                long current = 0;
                while (current < frames)
                {
                    int count = (int)Math.Min(buffer.Length / align, frames - current);
                    file.ReadExactly(buffer.AsSpan(0, count * align));
                    for (int i = 0; i < count; i++)
                    {
                        bool silent = true;
                        for (int ch = 0; ch < channels; ch++)
                        {
                            short sample = BinaryPrimitives.ReadInt16LittleEndian(buffer.AsSpan(i * align + ch * 2, 2));
                            int amplitude = Math.Abs((int)sample);
                            peak = Math.Max(peak, amplitude);
                            if (amplitude >= 32760) clipped++;
                            if (amplitude > 104) silent = false; // about -50 dBFS
                            double normalized = sample / 32768.0;
                            sumSquares[ch] += normalized * normalized;
                        }
                        long frame = current + i;
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
                if (clipped > frames * channels * 0.001)
                    issues.Add(new("AUDIO_CLIPPING", "error", null, null, "フルスケール付近のサンプルが継続しています"));
                var rmsByChannel = sumSquares.Select(sum => Math.Sqrt(sum / frames)).ToArray();
                if (channels == 2 && rmsByChannel.Max() > 0.01 &&
                    rmsByChannel.Min() * 4 < rmsByChannel.Max())
                    issues.Add(new("CHANNEL_IMBALANCE", "warning", null, null, "左右チャンネルの RMS に大きな差があります"));
                return new(true, !issues.Any(i => i.Severity == "error"), null, null, rate, channels,
                    (double)frames / rate, (double)peak / 32768,
                    Math.Sqrt(sumSquares.Sum() / (frames * channels)), issues.ToArray());
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
            => new(false, false, code, message, null, null, null, null, null, Array.Empty<AudioQaIssue>());
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
