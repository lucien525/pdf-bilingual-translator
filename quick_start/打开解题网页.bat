@echo off
chcp 65001 >nul
REM ★ 修复：访问网址_解题器.txt 写在项目根目录，先读真实端口
cd /d "%~dp0.."

if exist "访问网址_解题器.txt" (
    for /f "tokens=*" %%a in ('findstr /r "http://127.0.0.1" 访问网址_解题器.txt') do (
        set "URL=%%a"
        goto :open
    )
)

REM 没找到就退回默认端口
set "URL=http://127.0.0.1:7870"

:open
start "" "%URL%"
exit
