# Windows導入受け入れ検証（Issue #22）

この手順は正式ReleaseからYMM4プラグインとPython MCPサーバーをインストールした**Windows実機**で実行します。GitHub Actionsのモック・パッケージビルドだけでは、YMM4でのダブルクリック導入は証明できません。

1. YMM4を終了し、正式Releaseの `.ymme` をダブルクリックしてインストールします。
2. 同じReleaseのサーバーZIPをすべて展開し、`mcp-server/install.cmd` を実行します。管理者権限は不要です。
3. YMM4を起動して「ツール」→「MCP連携サーバー」で自動起動・ポートを確認します。
4. 展開したZIPフォルダーを削除しても構いません。インストール済みの `%LOCALAPPDATA%\YMM4MCP\mcp-server\check_installation.cmd` を実行します。成功時は `success: true`、検証したClaudeプロファイル数、認証済みstatus、ポート、バージョンが表示されます。

`check_installation.cmd` は**読み取り専用**です。永続インストール先のPython・server.py・requirements.txt、通常版と初期化済みMicrosoft Store版のClaude Desktop設定、Pythonとserver.pyの絶対パス、設定されたループバックポートを検査します。最後に既存の `check_connection.check()` を呼び、正しいトークンでstatusが成功し、不正なトークンではHTTP 401になることを確認します。APIトークンやClaudeの環境変数の値は表示しません。

失敗した場合は、YMM4のMCPサーバーが起動しているか、Claude Desktop設定が古いZIP展開先を指していないか、ポートが一致しているかを確認してください。失敗時も設定を変更しません。テストは`python -m unittest -v test_check_installation`で実行でき、HTTPはモックです。

**未検証事項:** この検証は実機のYMM4を自動操作せず、`.ymme` のダブルクリック操作そのもの、正式Releaseのasset公開、Claude Desktop内での実際のMCPツール実行は証明しません。これらの実機確認と関連PRのマージが完了するまでIssue #22はOpenのままにします。

## MCP stdio の実接続検査

`check_installation.cmd` は認証付きHTTP statusに加え、インストール先のPythonを**実際に子プロセス起動**して、各Claude Desktopプロファイルの環境変数でMCP `initialize` と `tools/list` を実行します。`ymm4_interact`・`ymm4_preview`・`ymm4_advanced`・`ymm4_analyze_video` の4ツールが見えなければ失敗します。**編集ツールの呼び出しは行いません**。成功時の `mcp_stdio: ready` と `mcp_profiles_checked` はClaudeのプロセス起動・MCPハンドシェイクの確認であり、Claude DesktopのGUIから実際に利用できることの保証ではありません。失敗時にAPIキー・トークン・プロファイル環境変数を出力しません。

CIでは `python -m unittest -v test_check_mcp_stdio` により同じSDKを用いた実プロセスのMCP初期化を確認します。これはWindows/YMM4実機での配布・プラグイン読み込み確認とは区別してください。
