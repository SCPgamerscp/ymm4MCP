using System;

namespace YMM4McpPlugin
{
    /// <summary>
    /// Restrict shell-launched links from the plugin UI to its two documented
    /// HTTPS destinations. Never pass a caller-provided URL to the shell.
    /// </summary>
    internal static class RepositoryLinkPolicy
    {
        private const string Repository = "https://github.com/SCPgamerscp/ymm4MCP";
        private const string Releases = Repository + "/releases";

        internal static bool TryGetSafeUrl(Uri? requested, out string safeUrl)
        {
            safeUrl = string.Empty;
            if (requested is null || !requested.IsAbsoluteUri ||
                !string.Equals(requested.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(requested.IdnHost, "github.com", StringComparison.OrdinalIgnoreCase) ||
                requested.Port != 443 || requested.UserInfo.Length != 0 ||
                requested.Query.Length != 0 || requested.Fragment.Length != 0)
                return false;

            if (string.Equals(requested.AbsolutePath, "/SCPgamerscp/ymm4MCP", StringComparison.Ordinal))
                safeUrl = Repository;
            else if (string.Equals(requested.AbsolutePath, "/SCPgamerscp/ymm4MCP/releases", StringComparison.Ordinal))
                safeUrl = Releases;
            else
                return false;

            return true;
        }
    }
}
