@echo off
cd /d "%~dp0"
setlocal

REM ---- 包内便携 Python（无需用户安装，随包拷贝即可）----
set "PY=%~dp0..\runtime\python.exe"

if not exist "%PY%" (
    echo [ERROR] 未找到 runtime\python.exe，请确认整个 config-forge-portable 文件夹已完整拷贝。
    pause
    exit /b 1
)

REM ---- 端口（默认 18000，避开 Windows 动态端口范围 1024~15000；可传参：start.bat 9000）----
set "PORT=18000"
if not "%~1"=="" set "PORT=%~1"

echo ==============================================
echo  CONFIG FORGE Platform (便携版)
echo  Python : %PY%
echo  URL    : http://localhost:%PORT%
echo  Stop   : 关闭本窗口或按 Ctrl+C
echo ==============================================
echo.

REM ---- 3 秒后自动打开浏览器 ----
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:%PORT%"

REM ---- 在当前窗口运行服务 ----
"%PY%" app.py
pause
