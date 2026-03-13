#!/usr/bin/env python3
"""
MonitorLuna Agent - Windows 系统托盘后台服务
WebUI 配置界面 + WebSocket 连接到 Koishi
"""
import asyncio
import base64
import io
import json
import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

import psutil
import pyautogui
import pystray
import websockets
import win32gui
import win32process
from aiohttp import web
from PIL import Image, ImageDraw

# ── 配置 ──────────────────────────────────────────────────────────────────────
CONFIG_PATH = Path(__file__).parent / "config.json"
WEBUI_PORT = 6315

DEFAULT_CONFIG = {
    "url": "ws://127.0.0.1:5140/monitorluna",
    "token": "",
    "device_id": os.environ.get("COMPUTERNAME", "my-pc"),
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── GPU 信息 ──────────────────────────────────────────────────────────────────
def get_gpu_info() -> list:
    gpus = []
    try:
        import GPUtil
        for g in GPUtil.getGPUs():
            if any(k in g.name for k in ("Intel", "UHD", "HD Graphics")):
                continue
            gpus.append({
                "name": g.name,
                "load": round(g.load * 100, 1),
                "memory_used": round(g.memoryUsed),
                "memory_total": round(g.memoryTotal),
            })
        if gpus:
            return gpus
    except Exception:
        pass

    try:
        import wmi
        w = wmi.WMI()
        for gpu in w.Win32_VideoController():
            name = gpu.Name or ""
            if any(k in name for k in ("Intel", "UHD", "HD Graphics")):
                continue
            vram_total = round((gpu.AdapterRAM or 0) / 1024 / 1024)
            gpus.append({
                "name": name,
                "load": -1,
                "memory_used": -1,
                "memory_total": vram_total,
            })
    except Exception:
        pass
    return gpus


# ── 窗口信息 ──────────────────────────────────────────────────────────────────
def get_window_info() -> dict:
    hwnd = win32gui.GetForegroundWindow()
    title = win32gui.GetWindowText(hwnd)
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        process_name = proc.name()
    except Exception:
        pid = -1
        process_name = "unknown"

    return {"title": title, "process": process_name, "pid": pid}


# ── 系统状态 ──────────────────────────────────────────────────────────────────
def get_system_status() -> dict:
    cpu = psutil.cpu_percent(interval=0.5)
    mem = psutil.virtual_memory()
    return {
        "cpu_percent": round(cpu, 1),
        "memory": {
            "total": round(mem.total / 1024 / 1024),
            "used": round(mem.used / 1024 / 1024),
            "percent": round(mem.percent, 1),
        },
        "gpu": get_gpu_info(),
    }


# ── 截图 ──────────────────────────────────────────────────────────────────────
def take_screenshot() -> str:
    img = pyautogui.screenshot()
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


def take_window_screenshot() -> str:
    hwnd = win32gui.GetForegroundWindow()
    rect = win32gui.GetWindowRect(hwnd)
    x, y, x2, y2 = rect
    img = pyautogui.screenshot(region=(x, y, x2 - x, y2 - y))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ── WebSocket Agent ───────────────────────────────────────────────────────────
class MonitorLunaAgent:
    def __init__(self):
        self.config = load_config()
        self.status = "未连接"
        self.running = True
        self._loop = None
        self._ws = None
        self._last_window = None

    def reload_config(self):
        self.config = load_config()

    async def _handle_command(self, msg: dict) -> dict:
        cmd = msg.get("cmd", "")
        cmd_id = msg.get("id", "")
        try:
            if cmd == "screenshot":
                data = await asyncio.get_event_loop().run_in_executor(None, take_screenshot)
                return {"type": "result", "id": cmd_id, "ok": True, "data": data}
            elif cmd == "window_screenshot":
                data = await asyncio.get_event_loop().run_in_executor(None, take_window_screenshot)
                return {"type": "result", "id": cmd_id, "ok": True, "data": data}
            elif cmd == "window_info":
                data = await asyncio.get_event_loop().run_in_executor(None, get_window_info)
                return {"type": "result", "id": cmd_id, "ok": True, "data": json.dumps(data, ensure_ascii=False)}
            elif cmd == "system_status":
                data = await asyncio.get_event_loop().run_in_executor(None, get_system_status)
                return {"type": "result", "id": cmd_id, "ok": True, "data": json.dumps(data, ensure_ascii=False)}
            else:
                return {"type": "result", "id": cmd_id, "ok": False, "error": f"unknown command: {cmd}"}
        except Exception as e:
            return {"type": "result", "id": cmd_id, "ok": False, "error": str(e)}

    async def _run_once(self):
        cfg = self.config
        url = cfg["url"]
        self.status = f"连接中... {url}"
        try:
            async with websockets.connect(url, open_timeout=10, ping_interval=30, ping_timeout=10) as ws:
                self._ws = ws
                await ws.send(json.dumps({"type": "hello", "token": cfg["token"], "device_id": cfg["device_id"]}))
                ack = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
                if ack.get("type") != "hello_ack":
                    self.status = f"握手失败: {ack.get('message', '未知错误')}"
                    return
                self.status = f"已连接 ✓ ({cfg['device_id']})"
                self._last_window = None
                monitor_task = asyncio.get_event_loop().create_task(self._window_monitor(ws, cfg["device_id"]))
                try:
                    async for raw in ws:
                        if not self.running:
                            break
                        try:
                            msg = json.loads(raw)
                        except Exception:
                            continue
                        if msg.get("type") == "command":
                            resp = await self._handle_command(msg)
                            await ws.send(json.dumps(resp, ensure_ascii=False))
                finally:
                    monitor_task.cancel()
                    self._ws = None
        except Exception as e:
            self.status = f"断开: {e}"
            self._ws = None

    async def _window_monitor(self, ws, device_id: str):
        while True:
            try:
                info = await asyncio.get_event_loop().run_in_executor(None, get_window_info)
                key = (info["process"], info["title"])
                if key != self._last_window:
                    self._last_window = key
                    await ws.send(json.dumps({
                        "type": "activity",
                        "device_id": device_id,
                        "process": info["process"],
                        "title": info["title"],
                        "pid": info["pid"],
                    }, ensure_ascii=False))
            except Exception:
                pass
            await asyncio.sleep(2)

    async def run_forever(self):
        delay = 3
        while self.running:
            await self._run_once()
            if not self.running:
                break
            self.status = f"重连中... ({delay}s 后)"
            await asyncio.sleep(delay)
            delay = min(delay * 2, 60)

    def start_in_thread(self):
        def _thread():
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            self._loop.run_until_complete(self.run_forever())
        t = threading.Thread(target=_thread, daemon=True)
        t.start()

    def stop(self):
        self.running = False
        if self._loop:
            self._loop.call_soon_threadsafe(self._loop.stop)


# ── WebUI ─────────────────────────────────────────────────────────────────────
HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>MonitorLuna 设置</title>
<style>
body{font-family:sans-serif;max-width:600px;margin:40px auto;padding:20px;background:#f5f5f5}
.card{background:#fff;padding:24px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}
h1{margin:0 0 20px;font-size:24px;color:#333}
.status{padding:12px;background:#e3f2fd;border-radius:4px;margin-bottom:20px;color:#1976d2}
label{display:block;margin:16px 0 6px;font-weight:600;color:#555}
input{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}
button{background:#1976d2;color:#fff;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;font-size:14px;margin-top:20px}
button:hover{background:#1565c0}
.msg{padding:12px;margin-top:16px;border-radius:4px;display:none}
.msg.success{background:#c8e6c9;color:#2e7d32}
.msg.error{background:#ffcdd2;color:#c62828}
</style>
</head>
<body>
<div class="card">
<h1>MonitorLuna 设置</h1>
<div class="status" id="status">状态: 加载中...</div>
<form id="form">
<label>Koishi WebSocket URL</label>
<input type="text" id="url" placeholder="ws://127.0.0.1:5140/monitorluna" required>
<label>Token</label>
<input type="password" id="token" placeholder="与 Koishi 配置一致" required>
<label>Device ID</label>
<input type="text" id="device_id" placeholder="my-pc" required>
<button type="submit">保存并重连</button>
</form>
<div class="msg" id="msg"></div>
</div>
<script>
async function load(){
  const r=await fetch('/api/config');
  const d=await r.json();
  document.getElementById('url').value=d.url;
  document.getElementById('token').value=d.token;
  document.getElementById('device_id').value=d.device_id;
  updateStatus();
}
async function updateStatus(){
  const r=await fetch('/api/status');
  const d=await r.json();
  document.getElementById('status').textContent='状态: '+d.status;
}
document.getElementById('form').onsubmit=async(e)=>{
  e.preventDefault();
  const data={
    url:document.getElementById('url').value,
    token:document.getElementById('token').value,
    device_id:document.getElementById('device_id').value
  };
  const r=await fetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const msg=document.getElementById('msg');
  if(r.ok){
    msg.className='msg success';
    msg.textContent='保存成功，正在重新连接...';
    msg.style.display='block';
    setTimeout(updateStatus,2000);
  }else{
    msg.className='msg error';
    msg.textContent='保存失败';
    msg.style.display='block';
  }
};
load();
setInterval(updateStatus,3000);
</script>
</body>
</html>
"""


async def handle_index(request):
    return web.Response(text=HTML_TEMPLATE, content_type="text/html")


async def handle_get_config(request):
    agent = request.app["agent"]
    return web.json_response(agent.config)


async def handle_post_config(request):
    agent = request.app["agent"]
    data = await request.json()
    save_config(data)
    agent.stop()
    await asyncio.sleep(0.5)
    agent.config = data
    agent.running = True
    agent.start_in_thread()
    return web.json_response({"ok": True})


async def handle_status(request):
    agent = request.app["agent"]
    return web.json_response({"status": agent.status})


def start_webui(agent: MonitorLunaAgent):
    app = web.Application()
    app["agent"] = agent
    app.router.add_get("/", handle_index)
    app.router.add_get("/api/config", handle_get_config)
    app.router.add_post("/api/config", handle_post_config)
    app.router.add_get("/api/status", handle_status)
    web.run_app(app, host="127.0.0.1", port=WEBUI_PORT, print=None)


# ── 托盘图标 ──────────────────────────────────────────────────────────────────
def create_tray_icon() -> Image.Image:
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([4, 4, size - 4, size - 4], fill=(30, 144, 255))
    draw.rectangle([18, 24, 46, 40], fill="white")
    draw.polygon([(24, 40), (40, 40), (32, 52)], fill="white")
    return img


def main():
    agent = MonitorLunaAgent()
    agent.start_in_thread()

    # WebUI 在后台线程运行
    threading.Thread(target=start_webui, args=(agent,), daemon=True).start()

    def on_settings(icon, item):
        webbrowser.open(f"http://127.0.0.1:{WEBUI_PORT}")

    def on_quit(icon, item):
        agent.stop()
        icon.stop()

    menu = pystray.Menu(
        pystray.MenuItem("打开设置", on_settings),
        pystray.MenuItem("退出", on_quit),
    )

    icon = pystray.Icon("monitorluna", create_tray_icon(), "MonitorLuna Agent", menu)

    def update_tooltip():
        while agent.running:
            try:
                icon.title = f"MonitorLuna - {agent.status}"
            except Exception:
                pass
            time.sleep(3)

    threading.Thread(target=update_tooltip, daemon=True).start()
    icon.run()


if __name__ == "__main__":
    main()
