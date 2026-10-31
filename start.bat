@echo off
chcp 65001 >nul
rem ============================================
rem  无痛实验报告生成器 · 一键启动
rem  双击本文件即可。依赖装在 libs/，无需全局安装。
rem ============================================
setlocal
cd /d "%~dp0"

set "PYTHON="
where py >nul 2>nul && set "PYTHON=py -3"
if not defined PYTHON (
  where python >nul 2>nul && set "PYTHON=python"
)
if not defined PYTHON (
  echo [错误] 未检测到 Python3，请先安装：https://www.python.org/downloads/
  echo        安装时勾选 "Add Python to PATH"。
  pause
  exit /b 1
)

echo [启动] 正在启动，首次可能稍慢，请稍候…
%PYTHON% "scripts\launcher.py"
if errorlevel 1 (
  echo.
  echo [提示] 若为依赖缺失，请先运行： python install_deps.py
  pause
)
endlocal