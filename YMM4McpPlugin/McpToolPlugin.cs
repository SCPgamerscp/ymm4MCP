using System;
using YukkuriMovieMaker.Plugin;

namespace YMM4McpPlugin
{
    /// <summary>
    /// YMM4 MCP連携ツールプラグイン
    /// AIがMCP経由でYMM4を操作できるようにするHTTPサーバーを起動します
    /// </summary>
    [PluginDetails(AuthorName = "SCPgamerscp")]
    public class McpToolPlugin : IToolPlugin
    {
        public McpToolPlugin() => McpViewModel.InitializePlugin();

        public string Name => "MCP連携サーバー";
        public Type ViewModelType => typeof(McpViewModel);
        public Type ViewType => typeof(McpView);
    }
}
