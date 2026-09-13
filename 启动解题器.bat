@echo off
chcp 65001 >nul
cd /d %~dp0

echo ============================================================
echo   PDF / Word / PPT 作业解题器 启动器
echo ============================================================
echo.

REM ============ 找 conda（跟 重启翻译器.bat 一致：找到第一个就跳出）============
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

if not defined CONDA_ACTIVATE (
    echo [错误] 未找到 conda 的 activate.bat
    echo.
    echo 请参照 重启翻译器.bat 的说明手动填路径。
    pause
    exit /b 1
)

echo [信息] 找到 conda：%CONDA_ACTIVATE%
echo.

call "%CONDA_ACTIVATE%" trans

if errorlevel 1 (
    echo.
    echo [错误] 激活 conda 环境 trans 失败
    echo   1. 检查环境名：conda env list
    echo   2. 若非 trans，把本文件里的 trans 改成你的环境名
    echo.
    pause
    exit /b 1
)

echo [信息] 检查 matplotlib / numpy……
python -c "import matplotlib, numpy" 2>nul
if errorlevel 1 (
    echo [信息] 缺失依赖，开始安装（约 30 秒）……
    python -m pip install matplotlib numpy -i https://pypi.tuna.tsinghua.edu.cn/simple
    echo.
)

echo [信息] 启动解题器……
echo.

python homework_app.py

if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出，退出码：%errorlevel%
    echo 请检查上方错误信息。
    echo.
)

pause