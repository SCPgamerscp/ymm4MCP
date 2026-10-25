using YMM4McpPlugin;

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
Console.WriteLine("Added item verification tests passed");

public class VideoItem { }
public sealed class CustomVideoItem : VideoItem { }
public sealed class NotVideoItem { }
