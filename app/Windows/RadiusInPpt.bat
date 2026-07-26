@echo off
REM ============================================================
REM RadiusInPpt -- Windows launcher (Office Add-in route)
REM
REM Equivalent of the macOS bash launcher (app/MacOS/RadiusInPpt):
REM   1. Find node (try `where node` first, fallback paths)
REM   2. Start static file server in background (tools/serve.js)
REM   3. Copy manifest.xml to %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
REM   4. Show popup: fully quit PowerPoint then reopen
REM
REM i18n: popup text follows Windows display language (zh / en).
REM Requires Windows 10+ (PowerShell + .NET 3.5/4 for MessageBox).
REM ============================================================
setlocal

REM --- 1. Paths --------------------------------------------------------
set "APP_DIR=%~dp0..\.."
set "RES_DIR=%APP_DIR%\Resources"
set "MANIFEST=%RES_DIR%\manifest.xml"
set "SERVER_JS=%RES_DIR%\tools\serve.js"
set "LOG_FILE=%TEMP%\radius_in_ppt.log"
set "PID_FILE=%TEMP%\radius_in_ppt.pid"
set "WEF_DIR=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

REM --- 2. Find node ----------------------------------------------------
set "NODE_BIN="
where node >nul 2>nul
if %ERRORLEVEL% == 0 (
    for /f "delims=" %%i in ('where node') do (
        set "NODE_BIN=%%i"
        goto :node_found
    )
)
if exist "C:\Program Files\nodejs\node.exe" set "NODE_BIN=C:\Program Files\nodejs\node.exe"
if exist "C:\Program Files (x86)\nodejs\node.exe" set "NODE_BIN=C:\Program Files (x86)\nodejs\node.exe"
if not defined NODE_BIN (
    powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Node.js (18+) not found.`n`nRadiusInPpt needs Node.js to run the local file server.`n`nInstall from https://nodejs.org/ (LTS recommended).', 'RadiusInPpt', 'OK', 'Warning')"
    start https://nodejs.org/
    exit /b 1
)
:node_found
echo [RadiusInPpt] using node: %NODE_BIN%

REM --- 3. Start server (skip if port 3000 already listening) ---------
netstat -ano | findstr ":3000" | findstr "LISTENING" >nul
if %ERRORLEVEL% == 0 (
    echo [RadiusInPpt] server already running on port 3000
    goto :server_ready
)

cd /d "%RES_DIR%"
start "RadiusInPpt Server" /b "" "%NODE_BIN%" "%SERVER_JS%" > "%LOG_FILE%" 2>&1
REM Capture the spawned PID via start /WAIT would block, so we just sleep + verify port
ping -n 3 127.0.0.1 >nul
netstat -ano | findstr ":3000" | findstr "LISTENING" >nul
if not %ERRORLEVEL% == 0 (
    powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Failed to start the server. See %LOG_FILE% for details.', 'RadiusInPpt', 'OK', 'Error')"
    type "%LOG_FILE%" 2>nul
    exit /b 1
)
echo [RadiusInPpt] server started OK

:server_ready

REM --- 4. Copy manifest to wef ---------------------------------------
if not exist "%WEF_DIR%" mkdir "%WEF_DIR%"
copy /Y "%MANIFEST%" "%WEF_DIR%\manifest.xml" >nul
if not %ERRORLEVEL% == 0 (
    powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Failed to copy manifest.xml to %WEF_DIR%', 'RadiusInPpt', 'OK', 'Error')"
    exit /b 1
)
echo [RadiusInPpt] manifest copied to %WEF_DIR%

REM --- 5. Show popup (zh / en based on system locale) ----------------
REM Use PowerShell to detect culture and show a localized MessageBox.
set "PS_SCRIPT=%TEMP%\radius_in_ppt_popup.ps1"
> "%PS_SCRIPT%" echo $culture = [System.Globalization.CultureInfo]::CurrentCulture
>> "%PS_SCRIPT%" echo $lang = $culture.TwoLetterISOLanguageName
>> "%PS_SCRIPT%" echo Add-Type -AssemblyName PresentationFramework
>> "%PS_SCRIPT%" echo if ($lang -eq 'zh') {
>> "%PS_SCRIPT%" echo     $title = 'R 角调整'
>> "%PS_SCRIPT%" echo     if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
>> "%PS_SCRIPT%" echo         $msg = "R 角调整 已就位`n`n静态 server 已启动（http://localhost:3000）`nmanifest 已复制到 wef 目录。`n`n你需要**完全退出 PowerPoint**（文件 ^> 退出，或在任务管理器里结束 POWERPNT.EXE）后重新打开，新加载项才会生效。`n`n之后每次使用只需双击本 .bat 即可。"
>> "%PS_SCRIPT%" echo     } else {
>> "%PS_SCRIPT%" echo         $msg = "R 角调整 已就位`n`n静态 server 已启动（http://localhost:3000）`nmanifest 已复制到 wef 目录。`n`n打开 PowerPoint 后，ribbon 会出现「R 角调整」tab。"
>> "%PS_SCRIPT%" echo     }
>> "%PS_SCRIPT%" echo } else {
>> "%PS_SCRIPT%" echo     $title = 'RadiusInPpt'
>> "%PS_SCRIPT%" echo     if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
>> "%PS_SCRIPT%" echo         $msg = "RadiusInPpt is ready.`n`nStatic server started (http://localhost:3000).`nManifest copied to wef directory.`n`n**Fully quit PowerPoint** (File ^> Quit, or end POWERPNT.EXE in Task Manager) and reopen it for the new add-in to load.`n`nAfter that, just double-click this .bat to start the server each time."
>> "%PS_SCRIPT%" echo     } else {
>> "%PS_SCRIPT%" echo         $msg = "RadiusInPpt is ready.`n`nStatic server started (http://localhost:3000).`nManifest copied to wef directory.`n`nOpen PowerPoint and you'll see the 'RadiusInPpt' tab on the ribbon."
>> "%PS_SCRIPT%" echo     }
>> "%PS_SCRIPT%" echo }
>> "%PS_SCRIPT%" echo [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Information')

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
del "%PS_SCRIPT%" 2>nul

endlocal
exit /b 0
