using System;
using YukkuriMovieMaker.Plugin;

namespace YMM4McpPlugin
{
    /// <summary>
    /// YMM4 MCP連携ツールプラグイン
    /// AIがMCP経由でYMM4を操作できるようにするHTTPサーバーを起動します
    /// </summary>
    public class McpToolPlugin : IToolPlugin
    {
        public string Name => "MCP連携サーバー";
        public Type ViewModelType => typeof(McpViewModel);
        public Type ViewType => typeof(McpView);
    }
}
