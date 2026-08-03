@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PYTHON_LAUNCHER="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_LAUNCHER=py -3"
if not defined PYTHON_LAUNCHER (
  where python >nul 2>nul
  if not errorlevel 1 set "PYTHON_LAUNCHER=python"
)

if not defined PYTHON_LAUNCHER (
  echo 找不到 Python。
  echo 請先到 https://www.python.org/downloads/ 安裝 Python 3.11 或 3.12。
  echo 安裝時請勾選 Add Python to PATH，完成後再執行本檔案。
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] 第一次使用：正在建立執行環境……
  %PYTHON_LAUNCHER% -m venv ".venv"
  if errorlevel 1 goto :failed
)

echo [2/3] 正在檢查及安裝必要套件……
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed
".venv\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto :failed

echo [3/3] 正在啟動網站……
echo 請保留這個視窗；關閉後網站也會停止。
".venv\Scripts\python.exe" "負面新聞整合網站.py"
if errorlevel 1 goto :failed
exit /b 0

:failed
echo.
echo 安裝或啟動失敗。
echo 請確認網路可連線、Proxy／防火牆允許 Python，並將畫面內容提供給系統管理員。
pause
exit /b 1
