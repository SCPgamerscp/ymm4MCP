@echo off
setlocal EnableExtensions
set "exit_code=1"
pushd "%~dp0" 2>nul
if errorlevel 1 goto :directory_error
if not exist ".venv\Scripts\python.exe" goto :python_error
if not exist "check_connection.py" goto :script_error

echo Checking the YMM4 plugin connection...
".venv\Scripts\python.exe" "check_connection.py"
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" goto :check_error
echo Connection verified.
goto :finish

:python_error
echo [ERROR] MCP server is not installed. Run install.cmd first.
goto :finish
:script_error
echo [ERROR] check_connection.py is missing. Extract the entire MCP server ZIP first.
goto :finish
:check_error
echo [ERROR] Connection check failed. Start YMM4 and verify the MCP plugin is running.
goto :finish
:directory_error
echo [ERROR] Cannot access the MCP server directory.
if not defined CI pause
exit /b 1

:finish
popd
if not defined CI pause
exit /b %exit_code%
