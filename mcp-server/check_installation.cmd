@echo off
setlocal EnableExtensions
if not defined LOCALAPPDATA goto :missing
set "server_dir=%LOCALAPPDATA%\YMM4MCP\mcp-server"
if not exist "%server_dir%\.venv\Scripts\python.exe" goto :missing
if not exist "%server_dir%\check_installation.py" goto :missing
"%server_dir%\.venv\Scripts\python.exe" "%server_dir%\check_installation.py"
exit /b %ERRORLEVEL%
:missing
echo [ERROR] Persistent MCP server installation not found. Run install.cmd first.
exit /b 1
