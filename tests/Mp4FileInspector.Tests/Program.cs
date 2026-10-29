using System.Buffers.Binary;
using System.Text;
using YMM4McpPlugin;

static byte[] Box(string type, params byte[][] parts)
{
    int length = 8 + parts.Sum(p => p.Length);
    var result = new byte[length];
    BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(0, 4), (uint)length);
    Encoding.ASCII.GetBytes(type).CopyTo(result, 4);
    int offset = 8;
    foreach (var part in parts) { part.CopyTo(result, offset); offset += part.Length; }
    return result;
}

static void U32(byte[] data, int offset, uint value)
    => BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(offset, 4), value);

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var ftyp = Box("ftyp", Encoding.ASCII.GetBytes("isom"), new byte[4]); // major_brand + minor_version
var mvhd = new byte[20];
U32(mvhd, 12, 1000); U32(mvhd, 16, 2000);
var tkhd = new byte[84];
U32(tkhd, 76, 1920u << 16); U32(tkhd, 80, 1080u << 16);
static byte[] Handler(string type)
{
    var hdlr = new byte[12];
    Encoding.ASCII.GetBytes(type).CopyTo(hdlr, 8);
    return Box("mdia", Box("hdlr", hdlr));
}
var mdhd = new byte[20];
U32(mdhd, 12, 30000);
var stts = new byte[16];
U32(stts, 4, 1); U32(stts, 8, 60); U32(stts, 12, 1000);
var timedMedia = Box("mdia", Box("hdlr", new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, (byte)'v', (byte)'i', (byte)'d', (byte)'e' }),
    Box("mdhd", mdhd), Box("minf", Box("stbl", Box("stts", stts))));
var video = Box("trak", Box("tkhd", tkhd), timedMedia);
var sample = new byte[28];
BinaryPrimitives.WriteUInt16BigEndian(sample.AsSpan(16, 2), 2);
U32(sample, 24, 48000u << 16);
var stsd = new byte[8];
U32(stsd, 4, 1);
var audio = Box("trak", Box("mdia", Box("hdlr", new byte[] {
    0, 0, 0, 0, 0, 0, 0, 0, (byte)'s', (byte)'o', (byte)'u', (byte)'n' }),
    Box("minf", Box("stbl", Box("stsd", stsd, Box("mp4a", sample))))));
var moov = Box("moov", Box("mvhd", mvhd), video, audio);
var mdat = Box("mdat", new byte[] { 1, 2, 3, 4 });
var path = Path.GetTempFileName();
try
{
    void Write(params byte[][] parts) => File.WriteAllBytes(path, parts.SelectMany(p => p).ToArray());
    Write(ftyp, mdat, moov); // moov at end is common for non-fast-start files.
    var valid = Mp4FileInspector.Inspect(path);
    Check(valid.Verified && valid.HasAudio && valid.Brand == "isom", "valid MP4 rejected");
    Check(valid.DurationSeconds == 2 && valid.Width == 1920 && valid.Height == 1080, "metadata mismatch");
    Check(valid.AverageFps == 30, "stts average FPS mismatch");
    Check(valid.AudioSampleRate == 48000 && valid.AudioChannels == 2, "mp4a audio metadata mismatch");

    Write(Box("ftyp", Encoding.ASCII.GetBytes("isom")), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "ftyp without minor_version accepted");
    Write(Box("ftyp", Encoding.ASCII.GetBytes("isom"), new byte[5]), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "ftyp with incomplete compatible brand accepted");

    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd), video, Box("trak", Handler("soun"))));
    var unknownAudio = Mp4FileInspector.Inspect(path);
    Check(unknownAudio.Verified && unknownAudio.HasAudio && unknownAudio.AudioSampleRate is null,
        "missing sample description should not invent audio metadata");

    var badStts = (byte[])stts.Clone();
    U32(badStts, 4, 2);
    var badVideo = Box("trak", Box("tkhd", tkhd), Box("mdia", Box("hdlr", new byte[] {
        0, 0, 0, 0, 0, 0, 0, 0, (byte)'v', (byte)'i', (byte)'d', (byte)'e' }),
        Box("mdhd", mdhd), Box("minf", Box("stbl", Box("stts", badStts)))));
    Write(ftyp, Box("moov", Box("mvhd", mvhd), badVideo), mdat);
    var unknownFps = Mp4FileInspector.Inspect(path);
    Check(unknownFps.Verified && unknownFps.AverageFps is null, "bad stts should leave FPS unknown");

    Write(ftyp, mdat, ftyp, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate ftyp accepted");
    Write(ftyp, mdat, moov, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate moov accepted");
    Write(Box("ftyp", new byte[] { 0, (byte)'s', (byte)'o', (byte)'m' }, new byte[4]), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "non-printable major brand accepted");
    Write(Box("ftyp", Encoding.ASCII.GetBytes("    "), new byte[4]), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "blank major brand accepted");

    Write(ftyp, mdat);
    Check(!Mp4FileInspector.Inspect(path).Verified, "missing moov accepted");
    Write(ftyp, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "missing mdat accepted");
    Write(ftyp, Box("moov", Box("mvhd", mvhd), audio), mdat);
    Check(!Mp4FileInspector.Inspect(path).Verified, "audio-only MP4 accepted");
    Write(ftyp, moov, mdat[..^1]);
    Check(!Mp4FileInspector.Inspect(path).Verified, "truncated box accepted");
    Write(ftyp, moov, mdat);
    File.WriteAllBytes(path, File.ReadAllBytes(path).AsSpan(0, 11).ToArray());
    Check(!Mp4FileInspector.Inspect(path).Verified, "truncated ftyp accepted");

    // Compatible brands are part of the verified container identity.
    Write(Box("ftyp", Encoding.ASCII.GetBytes("isom"), new byte[4], new byte[4]), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "non-printable compatible brand accepted");
    Write(Box("ftyp", Encoding.ASCII.GetBytes("isom"), new byte[4],
        Encoding.ASCII.GetBytes("    ")), mdat, moov);
    Check(!Mp4FileInspector.Inspect(path).Verified, "blank compatible brand accepted");
    Write(Box("ftyp", Encoding.ASCII.GetBytes("isom"), new byte[4],
        Encoding.ASCII.GetBytes("iso2")), mdat, moov);
    Check(Mp4FileInspector.Inspect(path).Verified, "valid compatible brand rejected");

    // Conflicting duplicate movie and track metadata must not be resolved by
    // silently choosing the first or last box.
    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd), Box("mvhd", mvhd), video));
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate mvhd accepted");
    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd),
        Box("trak", Box("tkhd", tkhd), Box("tkhd", tkhd), timedMedia)));
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate tkhd accepted");
    var videoHandler = new byte[12];
    Encoding.ASCII.GetBytes("vide").CopyTo(videoHandler, 8);
    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd),
        Box("trak", Box("tkhd", tkhd),
            Box("mdia", Box("hdlr", videoHandler), Box("hdlr", videoHandler)))));
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate hdlr accepted");
    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd),
        Box("trak", Box("tkhd", tkhd),
            Box("mdia", Box("hdlr", videoHandler), Box("mdhd", mdhd),
                Box("mdhd", mdhd)))));
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate mdhd accepted");
    Write(ftyp, mdat, Box("moov", Box("mvhd", mvhd),
        Box("trak", Box("tkhd", tkhd),
            Box("mdia", Box("hdlr", videoHandler), Box("minf"),
                Box("minf")))));
    Check(!Mp4FileInspector.Inspect(path).Verified, "duplicate minf accepted");

}
finally { File.Delete(path); }

Console.WriteLine("MP4 inspection tests passed");
