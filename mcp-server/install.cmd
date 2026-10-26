@echo off
setlocal EnableExtensions
set "exit_code=1"
pushd "%~dp0" 2>nul
if errorlevel 1 goto :directory_error
if not exist "setup.ps1" goto :script_error
where powershell.exe >nul 2>&1
if errorlevel 1 goto :powershell_error

echo Setting up the YMM4 MCP server and Claude Desktop integration...
rem ExecutionPolicy applies to this PowerShell process only; no administrator access is required.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CD%\setup.ps1" -Install -ConfigureClaudeDesktop
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" goto :setup_error

echo.
echo Setup completed. Restart Claude Desktop to load the YMM4 MCP server.
echo Start YMM4, then run check_connection.cmd to verify connectivity.
goto :finish

:script_error
echo [ERROR] setup.ps1 is missing. Extract the entire MCP server ZIP first.
goto :finish
:powershell_error
echo [ERROR] Windows PowerShell was not found.
goto :finish
:setup_error
echo [ERROR] Setup failed. Review the error above. Existing Claude settings are preserved on setup failure.
goto :finish
:directory_error
echo [ERROR] Cannot access the installer directory.
if not defined CI pause
exit /b 1

:finish
popd
if not defined CI pause
exit /b %exit_code%
