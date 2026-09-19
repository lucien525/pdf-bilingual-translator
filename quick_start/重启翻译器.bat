@echo off
chcp 65001 >nul
REM ★ 修复：bat 在 quick_start 子目录，退回项目根目录再启动 python
cd /d "%~dp0.."

echo ============================================================
echo   PDF / Word / PPT 翻译器 启动器
echo ============================================================
echo.

REM ============================================================
REM  自动探测 conda 安装位置（依次尝试常见路径）
REM  ★ 修复：找到第一个存在的就 goto 跳出循环，避免多环境时
REM         被列表靠后的路径覆盖。
REM ============================================================
set "CONDA_ACTIVATE="

for %%P in (
    "%USERPROFILE%\miniconda3\Scripts\activate.bat"
    "%USERPROFILE%\anaconda3\Scripts\activate.bat"
    "%USERPROFILE%\AppData\Local\miniconda3\Scripts\activate.bat"
    "%USERPROFILE%\AppData\Local\Continuum\anaconda3\Scripts\activate.bat"
    "C:\ProgramData\miniconda3\Scripts\activate.bat"
    "C:\ProgramData\Anaconda3\Scripts\activate.bat"
    "C:\miniconda3\Scripts\activate.bat"
    "C:\Anaconda3\Scripts\activate.bat"
    "D:\app\miniconda\conda\Scripts\activate.bat"
    "D:\miniconda3\Scripts\activate.bat"
    "D:\Anaconda3\Scripts\activate.bat"
    "E:\miniconda3\Scripts\activate.bat"
    "E:\Anaconda3\Scripts\activate.bat"
) do (
    if exist %%P (
        set "CONDA_ACTIVATE=%%~P"
        goto :found_conda
    )
)

:found_conda

REM ============================================================
REM  如果自动探测失败，尝试从 PATH 里找 conda
REM  ★ 修复：找到第一个就 break，避免多环境时被覆盖
REM ============================================================
if not defined CONDA_ACTIVATE (
    for /f "delims=" %%C in ('where conda 2^>nul') do (
        set "CONDA_EXE=%%C"
        goto :got_conda
    )
    :got_conda
    if defined CONDA_EXE (
        for %%D in ("%CONDA_EXE%") do set "CONDA_DIR=%%~dpD"
        if exist "%CONDA_DIR%activate.bat" (
            set "CONDA_ACTIVATE=%CONDA_DIR%activate.bat"
        )
    )
)

REM ============================================================
REM  还找不到？让用户手动填
REM ============================================================
if not defined CONDA_ACTIVATE (
    echo [错误] 未找到 conda 的 activate.bat
    echo.
    echo 请手动修改本文件，找到下面这一行：
    echo     set "CONDA_ACTIVATE=你的路径"
    echo.
    echo 如何找路径：
    echo   1. 打开命令行，执行：where conda
    echo   2. 得到类似 D:\app\miniconda\conda\Scripts\conda.exe
    echo   3. 把最后一段 conda.exe 换成 activate.bat，就是这个路径
    echo.
    pause
    exit /b 1
)

echo [信息] 找到 conda：%CONDA_ACTIVATE%
echo.

REM ============================================================
REM  激活环境并启动
REM ============================================================
call "%CONDA_ACTIVATE%" trans

if errorlevel 1 (
    echo.
    echo [错误] 激活 conda 环境 trans 失败
    echo 可能原因：
    echo   1. 环境名不是 trans（请检查：conda env list）
    echo   2. 环境创建失败或被删除
    echo.
    echo 如果想用别的环境名，请编辑本文件，把这一行：
    echo     call "%CONDA_ACTIVATE%" trans
    echo 中的 trans 改成你的环境名。
    echo.
    pause
    exit /b 1
)

echo [信息] conda 环境已激活，启动翻译器……
echo.

python main.py

if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出，退出码：%errorlevel%
    echo 请检查上方错误信息。
    echo.
)

pause