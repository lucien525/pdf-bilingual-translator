@echo off
chcp 65001 >nul
REM ★ 修复：访问网址.txt 写在项目根目录
cd /d "%~dp0.."

REM ============================================================
REM  优先从 访问网址.txt 里读真实端口（程序可能不在 7860）
REM ============================================================
if exist "访问网址.txt" (
    for /f "tokens=*" %%a in ('findstr /r "http://127.0.0.1" 访问网址.txt') do (
        set "URL=%%a"
        goto :open
    )
)

REM 没找到就退回默认端口
set "URL=http://127.0.0.1:7860"

:open
start "" "%URL%"
exit