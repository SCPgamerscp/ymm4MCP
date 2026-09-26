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

var ftyp = Box("ftyp", Encoding.ASCII.GetBytes("isom"));
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
var video = Box("trak", Box("tkhd", tkhd), Handler("vide"));
var audio = Box("trak", Handler("soun"));
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
}
finally { File.Delete(path); }

Console.WriteLine("MP4 inspection tests passed");
