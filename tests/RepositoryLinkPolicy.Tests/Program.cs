using System;
using YMM4McpPlugin;

static void Check(string input, bool allowed, string? canonical = null)
{
    if (!Uri.TryCreate(input, UriKind.RelativeOrAbsolute, out var uri))
        throw new Exception($"Invalid test URI: {input}");

    bool result = RepositoryLinkPolicy.TryGetSafeUrl(uri, out string safeUrl);
    if (result != allowed)
        throw new Exception($"Unexpected decision for {input}: {result}");
    if (allowed && safeUrl != canonical)
        throw new Exception($"Incorrect canonical URL for {input}: {safeUrl}");
    if (!allowed && safeUrl != string.Empty)
        throw new Exception($"Rejected URI retained an executable URL: {input}");
}

const string repo = "https://github.com/SCPgamerscp/ymm4MCP";
const string releases = repo + "/releases";
Check(repo, true, repo);
Check(releases, true, releases);
Check("https://GITHUB.COM/SCPgamerscp/ymm4MCP", true, repo);
Check("http://github.com/SCPgamerscp/ymm4MCP", false);
Check("file:///C:/Windows/System32/calc.exe", false);
Check("javascript:alert(1)", false);
Check("https://github.com.evil.example/SCPgamerscp/ymm4MCP", false);
Check("https://github.com@evil.example/SCPgamerscp/ymm4MCP", false);
Check("https://user@github.com/SCPgamerscp/ymm4MCP", false);
Check("https://github.com:444/SCPgamerscp/ymm4MCP", false);
Check(repo + "?redirect=https://evil.example", false);
Check(repo + "#fragment", false);
Check(repo + "/releases/latest", false);
Check(repo + "/../../evil", false);
Check(repo + "%2Freleases", false);
Check("/SCPgamerscp/ymm4MCP", false);
if (RepositoryLinkPolicy.TryGetSafeUrl(null, out var rejected) || rejected != string.Empty)
    throw new Exception("Null URL was not rejected");
Console.WriteLine("RepositoryLinkPolicy: 17 safety cases passed.");
