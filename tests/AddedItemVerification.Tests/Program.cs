using YMM4McpPlugin;
using System.Globalization;

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

var valid = new VideoItem();
var derived = new CustomVideoItem();
var unrelated = new NotVideoItem();
Check(AddedItemVerification.IsExpectedType(valid, "VideoItem"), "exact type rejected");
Check(AddedItemVerification.IsExpectedType(derived, "VideoItem"), "derived type rejected");
Check(!AddedItemVerification.IsExpectedType(unrelated, "VideoItem"),
      "substring/suffix spoofed item type");
Check(!AddedItemVerification.IsExpectedType(new object(), "VideoItem"),
      "unrelated type accepted");
Check(!AddedItemVerification.IsExpectedType(null, "VideoItem"), "null accepted");
Check(ReferenceEquals(AddedItemVerification.UniqueExpectedNewcomer(
      new object[] { unrelated, valid }, "VideoItem"), valid),
      "unique expected newcomer not selected");
Check(AddedItemVerification.UniqueExpectedNewcomer(
      new object[] { valid, derived }, "VideoItem") == null,
      "multiple matching newcomers silently selected");
Check(AddedItemVerification.UniqueExpectedNewcomer(
      new object[] { unrelated }, "VideoItem") == null,
      "mismatched newcomer silently selected");
Check(AddedItemVerification.UniqueExpectedNewcomer(
      Array.Empty<object>(), "VideoItem") == null,
      "empty newcomer list silently accepted");
Check(AddedItemVerification.LengthMatches(12, 12), "matching length rejected");
Check(AddedItemVerification.LengthMatches(12L, 12), "matching long length rejected");
Check(!AddedItemVerification.LengthMatches(13, 12), "wrong length accepted");
Check(!AddedItemVerification.LengthMatches("12", 12), "string length accepted");

var savedCulture = CultureInfo.CurrentCulture;
try
{
    CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
    Check(ItemSnapshotFormatting.Format(1.25m) == "1.25", "decimal revision is locale dependent");
    Check(ItemSnapshotFormatting.Format(2.5d) == "2.5", "floating revision is locale dependent");
    Check(ItemSnapshotFormatting.Format(1234L) == "1234", "integer revision changed");
    Check(ItemSnapshotFormatting.Format(null) == "", "null revision changed");
    Check(ItemSnapshotFormatting.Format("日本語") == "日本語", "string revision changed");
}
finally { CultureInfo.CurrentCulture = savedCulture; }
Console.WriteLine("Added item verification tests passed");

public class VideoItem { }
public sealed class CustomVideoItem : VideoItem { }
public sealed class NotVideoItem { }
