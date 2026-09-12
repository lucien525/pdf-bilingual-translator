@echo off
chcp 65001 >nul
cd /d %~dp0

REM ========== 改成你的 conda 安装路径（和 启动翻译器.bat 里一样）==========
set CONDA_ACTIVATE=D:\app\miniconda\conda\Scripts\activate.bat
REM ======================================================================

call "%CONDA_ACTIVATE%" trans
python bilingual_app.py
pause