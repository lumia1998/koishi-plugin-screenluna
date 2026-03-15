#!/usr/bin/env python3
"""
MonitorLuna Agent - 跨平台系统托盘后台服务
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
import platform

import psutil
import pyautogui
import websockets
from aiohttp import web
from PIL import Image, ImageDraw

# Windows 特定导入（可选）
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    try:
        from ctypes import windll, WINFUNCTYPE, c_int, c_void_p, byref
        from ctypes.wintypes import MSG
        import win32gui
        import win32process
        import win32api
        import win32con
        import pystray
        WINDOWS_FEATURES = True
    except ImportError:
        WINDOWS_FEATURES = False
        print("Warning: Windows-specific features disabled (missing pywin32)")
else:
    WINDOWS_FEATURES = False

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


# ── 输入统计 ──────────────────────────────────────────────────────────────────
# 全局输入统计数据
_input_stats_lock = threading.Lock()
_app_stats = {}  # process_name -> {display_name, key_presses, left_clicks, right_clicks, scroll_distance}
_icon_cache = {}  # process_name -> base64 string
_last_input_time = time.time()  # 最后一次输入时间（用于 AFK 检测）
AFK_THRESHOLD = 180  # 3 分钟无输入视为 AFK

if IS_WINDOWS and WINDOWS_FEATURES:
    # Windows 钩子常量
    WH_KEYBOARD_LL = 13
    WH_MOUSE_LL = 14
    WM_KEYDOWN = 0x0100
    WM_SYSKEYDOWN = 0x0104
    WM_KEYUP = 0x0101
    WM_SYSKEYUP = 0x0105
    WM_LBUTTONDOWN = 0x0201
    WM_RBUTTONDOWN = 0x0204
    WM_MOUSEWHEEL = 0x020A
    WHEEL_DELTA = 120

    _keyboard_hook = None
    _mouse_hook = None
    _hook_thread = None
    _pressed_keys = set()  # 跟踪已按下的键，防止长按重复计数

    def _get_current_process_name() -> str:
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name()
        except Exception:
            return "unknown"

    def _ensure_app_entry(process_name: str):
        if process_name not in _app_stats:
            _app_stats[process_name] = {
                "display_name": process_name.replace(".exe", ""),
                "key_presses": 0,
                "left_clicks": 0,
                "right_clicks": 0,
                "scroll_distance": 0.0,
            }

    def _keyboard_proc(nCode, wParam, lParam):
        global _last_input_time
        if nCode >= 0:
            # 从 lParam 读取虚拟键码（KBDLLHOOKSTRUCT 第一个字段）
            import ctypes
            vk_code = ctypes.cast(lParam, ctypes.POINTER(ctypes.c_ulong))[0]
            if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                # 只在键第一次按下时计数，忽略长按重复
                if vk_code not in _pressed_keys:
                    _pressed_keys.add(vk_code)
                    _last_input_time = time.time()
                    process_name = _get_current_process_name()
                    with _input_stats_lock:
                        _ensure_app_entry(process_name)
                        _app_stats[process_name]["key_presses"] += 1
            elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                _pressed_keys.discard(vk_code)
        return windll.user32.CallNextHookEx(None, nCode, wParam, lParam)

    def _mouse_proc(nCode, wParam, lParam):
        global _last_input_time
        if nCode >= 0:
            if wParam in (WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEWHEEL):
                _last_input_time = time.time()
                process_name = _get_current_process_name()
                with _input_stats_lock:
                    _ensure_app_entry(process_name)
                    if wParam == WM_LBUTTONDOWN:
                        _app_stats[process_name]["left_clicks"] += 1
                    elif wParam == WM_RBUTTONDOWN:
                        _app_stats[process_name]["right_clicks"] += 1
                    elif wParam == WM_MOUSEWHEEL:
                        # 从 MSLLHOOKSTRUCT.mouseData 高位字取滚动 delta，除以 120 得格数
                        import ctypes
                        mouse_data = ctypes.cast(lParam, ctypes.POINTER(ctypes.c_ulong))[2]
                        delta = ctypes.c_short(mouse_data >> 16).value
                        ticks = abs(delta) / WHEEL_DELTA
                        _app_stats[process_name]["scroll_distance"] += ticks
        return windll.user32.CallNextHookEx(None, nCode, wParam, lParam)

    KEYBOARD_PROC_TYPE = WINFUNCTYPE(c_int, c_int, c_void_p, c_void_p)
    MOUSE_PROC_TYPE = WINFUNCTYPE(c_int, c_int, c_void_p, c_void_p)
    _kb_proc_ref = KEYBOARD_PROC_TYPE(_keyboard_proc)
    _ms_proc_ref = MOUSE_PROC_TYPE(_mouse_proc)

    def _start_hooks():
        global _keyboard_hook, _mouse_hook
        try:
            _keyboard_hook = windll.user32.SetWindowsHookExW(WH_KEYBOARD_LL, _kb_proc_ref, None, 0)
            _mouse_hook = windll.user32.SetWindowsHookExW(WH_MOUSE_LL, _ms_proc_ref, None, 0)
            msg = MSG()
            while windll.user32.GetMessageW(byref(msg), None, 0, 0) != 0:
                windll.user32.TranslateMessage(byref(msg))
                windll.user32.DispatchMessageW(byref(msg))
        except Exception:
            pass
        finally:
            if _keyboard_hook:
                windll.user32.UnhookWindowsHookEx(_keyboard_hook)
            if _mouse_hook:
                windll.user32.UnhookWindowsHookEx(_mouse_hook)

    def start_input_monitoring():
        global _hook_thread
        _hook_thread = threading.Thread(target=_start_hooks, daemon=True)
        _hook_thread.start()

else:
    def start_input_monitoring():
        pass  # 非 Windows 平台不支持输入监听


# ── 图标提取 ──────────────────────────────────────────────────────────────────
def extract_icon_base64(process_name: str) -> str:
    if not (IS_WINDOWS and WINDOWS_FEATURES):
        return ""  # 非 Windows 平台不支持图标提取

    if process_name in _icon_cache:
        return _icon_cache[process_name]
    try:
        # 找到进程路径（优先从运行中的进程获取，参考 keyStats AppIconHelper）
        exe_path = None
        for proc in psutil.process_iter(["name", "exe"]):
            try:
                if proc.info["name"] == process_name and proc.info["exe"]:
                    exe_path = proc.info["exe"]
                    break
            except Exception:
                continue

        if not exe_path or not os.path.exists(exe_path):
            _icon_cache[process_name] = ""
            return ""

        # 用 win32api.ExtractIconEx 提取图标
        try:
            large, small = win32api.ExtractIconEx(exe_path, 0)
        except Exception:
            _icon_cache[process_name] = ""
            return ""

        icon_handle = None
        handles_to_destroy = []
        if large:
            icon_handle = large[0]
            handles_to_destroy.extend(large[1:])
            handles_to_destroy.extend(small)
        elif small:
            icon_handle = small[0]
            handles_to_destroy.extend(small[1:])

        if not icon_handle:
            _icon_cache[process_name] = ""
            return ""

        try:
            import win32ui
            # 创建 32x32 内存 DC 并绘制图标
            hdc_screen = win32gui.GetDC(0)
            hdc = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem = hdc.CreateCompatibleDC()
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, 32, 32)
            hdc_mem.SelectObject(hbmp)
            # 白色背景
            hdc_mem.FillSolidRect((0, 0, 32, 32), 0xFFFFFF)
            win32gui.DrawIconEx(hdc_mem.GetHandleOutput(), 0, 0, icon_handle, 32, 32, 0, None, win32con.DI_NORMAL)

            bmp_info = hbmp.GetInfo()
            bmp_bits = hbmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGBA",
                (bmp_info["bmWidth"], bmp_info["bmHeight"]),
                bmp_bits, "raw", "BGRA", 0, 1
            )
            img = img.resize((20, 20), Image.LANCZOS)

            hdc_mem.DeleteDC()
            hdc.DeleteDC()
            win32gui.ReleaseDC(0, hdc_screen)
            hbmp.DeleteObject()
        except Exception:
            img = None
        finally:
            win32gui.DestroyIcon(icon_handle)
            for h in handles_to_destroy:
                try:
                    win32gui.DestroyIcon(h)
                except Exception:
                    pass

        if img is None:
            _icon_cache[process_name] = ""
            return ""

        buf = io.BytesIO()
        img.save(buf, "PNG")
        result = base64.b64encode(buf.getvalue()).decode()
        _icon_cache[process_name] = result
        return result
    except Exception:
        _icon_cache[process_name] = ""
        return ""


def is_afk() -> bool:
    """检查当前是否处于 AFK 状态（超过 AFK_THRESHOLD 秒无输入）"""
    return (time.time() - _last_input_time) > AFK_THRESHOLD


def get_input_stats_snapshot() -> dict:
    """返回当前输入统计快照，并为每个应用添加图标"""
    with _input_stats_lock:
        snapshot = {k: dict(v) for k, v in _app_stats.items()}
    result = {}
    for process_name, stats in snapshot.items():
        icon = extract_icon_base64(process_name)
        result[process_name] = {
            "display_name": stats["display_name"],
            "icon_base64": icon,
            "key_presses": stats["key_presses"],
            "left_clicks": stats["left_clicks"],
            "right_clicks": stats["right_clicks"],
            "scroll_distance": stats["scroll_distance"],
        }
    return result


# ── 窗口信息 ──────────────────────────────────────────────────────────────────
def get_window_info() -> dict:
    if IS_WINDOWS and WINDOWS_FEATURES:
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            process_name = proc.name()
            return {"title": title, "process": process_name, "pid": pid}
        except Exception:
            return {"title": "", "process": "unknown", "pid": -1}
    else:
        # 非 Windows 平台简化实现
        return {"title": "N/A", "process": "unknown", "pid": -1}


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
    if IS_WINDOWS and WINDOWS_FEATURES:
        hwnd = win32gui.GetForegroundWindow()
        rect = win32gui.GetWindowRect(hwnd)
        x, y, x2, y2 = rect
        img = pyautogui.screenshot(region=(x, y, x2 - x, y2 - y))
    else:
        # 非 Windows 平台退回全屏截图
        img = pyautogui.screenshot()
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ── WebSocket Agent ───────────────────────────────────────────────────────────
class MonitorLunaAgent:
    def __init__(self):
        self.config = load_config()
        self.status = "未连接"
        self.running = True
        self.paused = False
        self._loop = None
        self._ws = None
        self._last_window = None
        start_input_monitoring()  # 启动输入监听

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
                stats_task = asyncio.get_event_loop().create_task(self._input_stats_sender(ws, cfg["device_id"]))
                status_task = asyncio.get_event_loop().create_task(self._update_status(cfg["device_id"]))
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
                    stats_task.cancel()
                    status_task.cancel()
                    self._ws = None
        except Exception as e:
            self.status = f"断开: {e}"
            self._ws = None

    async def _window_monitor(self, ws, device_id: str):
        last_title = None
        while True:
            try:
                if not self.paused:
                    info = await asyncio.get_event_loop().run_in_executor(None, get_window_info)
                    afk = await asyncio.get_event_loop().run_in_executor(None, is_afk)
                    key = (info["process"], info["title"])

                    # 窗口标题变化也算活跃（视频播放、直播等场景）
                    if info["title"] != last_title and info["title"]:
                        import builtins
                        # 直接写全局变量
                        globals()["_last_input_time"] = time.time()
                        last_title = info["title"]

                    # Send on window change OR always (to keep heartbeat endTime updated)
                    if key != self._last_window:
                        self._last_window = key
                    await ws.send(json.dumps({
                        "type": "activity",
                        "device_id": device_id,
                        "process": info["process"],
                        "title": info["title"],
                        "pid": info["pid"],
                        "afk": afk,
                    }, ensure_ascii=False))
            except Exception:
                pass
            await asyncio.sleep(2)

    async def _input_stats_sender(self, ws, device_id: str):
        while True:
            await asyncio.sleep(30)  # 每 30 秒发送一次
            try:
                if not self.paused:
                    stats = await asyncio.get_event_loop().run_in_executor(None, get_input_stats_snapshot)
                    if stats:
                        await ws.send(json.dumps({
                            "type": "input_stats",
                            "device_id": device_id,
                            "stats": stats,
                        }, ensure_ascii=False))
            except Exception:
                pass

    async def _update_status(self, device_id: str):
        while True:
            await asyncio.sleep(1)
            if self.paused:
                self.status = f"已暂停 ⏸ ({device_id})"
            else:
                self.status = f"已连接 ✓ ({device_id})"

    def toggle_pause(self):
        self.paused = not self.paused
        return self.paused

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
button.pause{background:#f57c00}
button.pause:hover{background:#e65100}
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
<button type="button" class="pause" id="pauseBtn" onclick="togglePause()">暂停监控</button>
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
  const paused=d.status.includes('暂停');
  const btn=document.getElementById('pauseBtn');
  btn.textContent=paused?'恢复监控':'暂停监控';
  btn.className=paused?'pause':'';
}
async function togglePause(){
  await fetch('/api/toggle_pause',{method:'POST'});
  updateStatus();
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


async def handle_toggle_pause(request):
    agent = request.app["agent"]
    paused = agent.toggle_pause()
    return web.json_response({"paused": paused})


def start_webui(agent: MonitorLunaAgent):
    app = web.Application()
    app["agent"] = agent
    app.router.add_get("/", handle_index)
    app.router.add_get("/api/config", handle_get_config)
    app.router.add_post("/api/config", handle_post_config)
    app.router.add_get("/api/status", handle_status)
    app.router.add_post("/api/toggle_pause", handle_toggle_pause)
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

    if IS_WINDOWS and WINDOWS_FEATURES:
        # Windows 托盘图标
        def on_settings(icon, item):
            webbrowser.open(f"http://127.0.0.1:{WEBUI_PORT}")

        def on_toggle_pause(icon, item):
            agent.toggle_pause()

        def on_quit(icon, item):
            agent.stop()
            icon.stop()

        menu = pystray.Menu(
            pystray.MenuItem("打开设置", on_settings),
            pystray.MenuItem("暂停/恢复", on_toggle_pause),
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
    else:
        # 非 Windows 平台：无托盘图标，保持运行
        print(f"MonitorLuna Agent 已启动")
        print(f"WebUI: http://127.0.0.1:{WEBUI_PORT}")
        print("按 Ctrl+C 退出")
        try:
            while agent.running:
                time.sleep(1)
        except KeyboardInterrupt:
            agent.stop()
            print("\n已退出")


if __name__ == "__main__":
    main()
