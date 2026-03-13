@echo off
chcp 65001 >nul

pythonw --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo [错误] 未检测到 Python 环境
        echo 请访问 https://www.python.org/downloads/ 下载安装 Python
        pause
        exit /b 1
    )
)

echo 检查依赖...
python -c "import websockets, pyautogui, PIL, psutil, pystray, win32gui, aiohttp" >nul 2>&1
if errorlevel 1 (
    echo 安装依赖中...
    pip install aiohttp websockets pyautogui pillow psutil GPUtil pywin32 pystray
)

start "" pythonw screenshot-server.py
exit
