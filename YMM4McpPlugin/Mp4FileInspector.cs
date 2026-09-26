using System;
using System.Buffers.Binary;
using System.IO;
using System.Text;

namespace YMM4McpPlugin
{
    internal sealed record Mp4Inspection(bool Verified, string Error, string? Brand = null,
        double? DurationSeconds = null, int? Width = null, int? Height = null, bool HasAudio = false);

    internal static class Mp4FileInspector
    {
        private readonly record struct Box(string Type, long End, long PayloadLength);

        public static Mp4Inspection Inspect(string path)
        {
            try
            {
                using var file = File.OpenRead(path);
                string? brand = null;
                bool moov = false, mdat = false, video = false, audio = false;
                double? duration = null;
                int? width = null, height = null;
                int boxes = 0;
                while (file.Position < file.Length)
                {
                    var box = ReadBox(file, file.Length);
                    if (++boxes > 20000) throw new InvalidDataException("Too many MP4 boxes");
                    if (boxes == 1 && box.Type != "ftyp") throw new InvalidDataException("ftyp is not the first box");
                    if (box.Type == "ftyp")
                    {
                        if (box.PayloadLength < 4) throw new InvalidDataException("ftyp is incomplete");
                        brand = Encoding.ASCII.GetString(Read(file, 4));
                    }
                    else if (box.Type == "mdat") mdat |= box.PayloadLength > 0;
                    else if (box.Type == "moov")
                    {
                        moov = true;
                        ReadMovie(file, box.End, ref duration, ref video, ref audio, ref width, ref height);
                    }
                    file.Position = box.End;
                }
                if (brand == null || !moov || !mdat || !video || duration is not > 0 || width is not > 0 || height is not > 0)
                    return new(false, "MP4 の映像トラック・尺・解像度・moov/mdat を確認できません");
                return new(true, "", brand, duration, width, height, audio);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or OverflowException or ArgumentException)
            {
                return new(false, "MP4 構造が不正です: " + ex.Message);
            }
        }

        private static void ReadMovie(Stream file, long end, ref double? duration, ref bool video,
            ref bool audio, ref int? width, ref int? height)
        {
            while (file.Position < end)
            {
                var box = ReadBox(file, end);
                if (box.Type == "mvhd")
                {
                    var header = Read(file, box.PayloadLength >= 32 && PeekVersion(file) == 1 ? 32 : 20, box.PayloadLength);
                    uint timescale = U32(header, header[0] == 1 ? 20 : 12);
                    ulong ticks = header[0] == 1 ? U64(header, 24) : U32(header, 16);
                    if (timescale > 0 && ticks > 0) duration = (double)ticks / timescale;
                }
                else if (box.Type == "trak") ReadTrack(file, box.End, ref video, ref audio, ref width, ref height);
                file.Position = box.End;
            }
        }

        private static void ReadTrack(Stream file, long end, ref bool video, ref bool audio,
            ref int? width, ref int? height)
        {
            string? handler = null;
            int trackWidth = 0, trackHeight = 0;
            while (file.Position < end)
            {
                var box = ReadBox(file, end);
                if (box.Type == "tkhd")
                {
                    int length = PeekVersion(file) == 1 ? 96 : 84;
                    var header = Read(file, length, box.PayloadLength);
                    trackWidth = checked((int)(U32(header, length - 8) >> 16));
                    trackHeight = checked((int)(U32(header, length - 4) >> 16));
                }
                else if (box.Type == "mdia")
                {
                    while (file.Position < box.End)
                    {
                        var child = ReadBox(file, box.End);
                        if (child.Type == "hdlr")
                        {
                            var header = Read(file, 12, child.PayloadLength);
                            handler = Encoding.ASCII.GetString(header, 8, 4);
                        }
                        file.Position = child.End;
                    }
                }
                file.Position = box.End;
            }
            if (handler == "vide")
            {
                video = true;
                if (trackWidth > 0 && trackHeight > 0) { width = trackWidth; height = trackHeight; }
            }
            if (handler == "soun") audio = true;
        }

        private static Box ReadBox(Stream file, long parentEnd)
        {
            if (parentEnd - file.Position < 8) throw new InvalidDataException("Incomplete box header");
            var header = Read(file, 8);
            ulong size = U32(header, 0);
            string type = Encoding.ASCII.GetString(header, 4, 4);
            long headerLength = 8;
            if (size == 1)
            {
                if (parentEnd - file.Position < 8) throw new InvalidDataException("Incomplete extended box header");
                size = U64(Read(file, 8), 0);
                headerLength = 16;
            }
            else if (size == 0) size = (ulong)(parentEnd - (file.Position - headerLength));
            if (size < (ulong)headerLength || size - (ulong)headerLength > (ulong)(parentEnd - file.Position))
                throw new InvalidDataException("Box exceeds its parent");
            return new(type, file.Position + (long)(size - (ulong)headerLength), (long)(size - (ulong)headerLength));
        }

        private static byte PeekVersion(Stream file)
        {
            int version = file.ReadByte();
            if (version < 0) throw new EndOfStreamException();
            file.Position--;
            if (version > 1) throw new InvalidDataException("Unsupported box version");
            return (byte)version;
        }

        private static byte[] Read(Stream file, int length, long available = long.MaxValue)
        {
            if (available < length) throw new InvalidDataException("Incomplete box payload");
            var bytes = new byte[length];
            file.ReadExactly(bytes);
            return bytes;
        }

        private static uint U32(byte[] data, int offset) => BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(offset, 4));
        private static ulong U64(byte[] data, int offset) => BinaryPrimitives.ReadUInt64BigEndian(data.AsSpan(offset, 8));
    }
}
