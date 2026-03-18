#!/usr/bin/env python3
"""
MonitorLuna Agent - 跨平台系统托盘后台服务
WebUI 配置界面 + WebSocket 连接到 Koishi
"""
import asyncio
import base64
import hashlib
import io
import json
import os
import sqlite3
import sys
import threading
import time
import webbrowser
import contextlib
from pathlib import Path
import platform
from datetime import datetime, timedelta

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
DB_PATH = Path(__file__).parent / "monitorluna.db"
WEBUI_PORT = 6315

DEFAULT_CONFIG = {
    "url": "ws://127.0.0.1:5140/monitorluna",
    "token": "",
    "device_id": os.environ.get("COMPUTERNAME", "my-pc"),
    "screenshot_enabled": False,
    "screenshot_interval": 15,
    "browser_token": "",
}


def init_database():
    """初始化本地 SQLite 数据库"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # 活动记录表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            process TEXT NOT NULL,
            title TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 输入统计表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS input_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            process TEXT NOT NULL,
            display_name TEXT NOT NULL,
            icon_hash TEXT,
            key_presses INTEGER DEFAULT 0,
            left_clicks INTEGER DEFAULT 0,
            right_clicks INTEGER DEFAULT 0,
            scroll_distance REAL DEFAULT 0,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Icon 缓存表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS icons (
            hash TEXT PRIMARY KEY,
            data TEXT NOT NULL
        )
    """)

    # 浏览器活动表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS browser_activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            domain TEXT NOT NULL,
            seconds REAL DEFAULT 0,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 创建索引
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_timestamp ON activity(timestamp)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_input_stats_timestamp ON input_stats(timestamp)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_browser_activity_timestamp ON browser_activity(timestamp)")

    conn.commit()
    conn.close()


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

# 浏览器活动统计（由本地 /ws/browser 端点接收）
_browser_stats_lock = threading.Lock()
_browser_stats = {}  # domain -> seconds

if IS_WINDOWS and WINDOWS_FEATURES:
    import ctypes
    from ctypes import wintypes

    # 声明 CallNextHookEx 的参数类型，避免 64 位 Windows 上 lParam 溢出
    windll.user32.CallNextHookEx.restype = wintypes.LPARAM
    windll.user32.CallNextHookEx.argtypes = [
        wintypes.HHOOK,   # hhk
        c_int,            # nCode
        wintypes.WPARAM,  # wParam
        wintypes.LPARAM,  # lParam
    ]

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
        if nCode >= 0:
            # lParam 是 LPARAM（整数），转为指针后读取 KBDLLHOOKSTRUCT.vkCode
            ptr = ctypes.cast(lParam, ctypes.POINTER(ctypes.c_ulong))
            vk_code = ptr[0]
            if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                if vk_code not in _pressed_keys:
                    _pressed_keys.add(vk_code)
                    process_name = _get_current_process_name()
                    with _input_stats_lock:
                        _ensure_app_entry(process_name)
                        _app_stats[process_name]["key_presses"] += 1
            elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                _pressed_keys.discard(vk_code)
        return windll.user32.CallNextHookEx(None, nCode, wParam, lParam)

    def _mouse_proc(nCode, wParam, lParam):
        if nCode >= 0:
            if wParam in (WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEWHEEL):
                process_name = _get_current_process_name()
                with _input_stats_lock:
                    _ensure_app_entry(process_name)
                    if wParam == WM_LBUTTONDOWN:
                        _app_stats[process_name]["left_clicks"] += 1
                    elif wParam == WM_RBUTTONDOWN:
                        _app_stats[process_name]["right_clicks"] += 1
                    elif wParam == WM_MOUSEWHEEL:
                        # lParam 是 LPARAM（整数），转为指针读取 MSLLHOOKSTRUCT.mouseData
                        ptr = ctypes.cast(lParam, ctypes.POINTER(ctypes.c_ulong))
                        mouse_data = ptr[2]
                        delta = ctypes.c_short(mouse_data >> 16).value
                        ticks = abs(delta) / WHEEL_DELTA
                        _app_stats[process_name]["scroll_distance"] += ticks
        return windll.user32.CallNextHookEx(None, nCode, wParam, lParam)

    KEYBOARD_PROC_TYPE = WINFUNCTYPE(wintypes.LPARAM, c_int, wintypes.WPARAM, wintypes.LPARAM)
    MOUSE_PROC_TYPE = WINFUNCTYPE(wintypes.LPARAM, c_int, wintypes.WPARAM, wintypes.LPARAM)
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
def extract_icon_base64(process_name: str) -> tuple[str, str]:
    """提取图标并返回 (hash, base64_data)"""
    if not (IS_WINDOWS and WINDOWS_FEATURES):
        return "", ""  # 非 Windows 平台不支持图标提取

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
            _icon_cache[process_name] = ("", "")
            return "", ""

        # 用 win32api.ExtractIconEx 提取图标
        try:
            large, small = win32api.ExtractIconEx(exe_path, 0)
        except Exception:
            _icon_cache[process_name] = ("", "")
            return "", ""

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
            _icon_cache[process_name] = ("", "")
            return "", ""

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
            _icon_cache[process_name] = ("", "")
            return "", ""

        buf = io.BytesIO()
        img.save(buf, "PNG")
        icon_data = base64.b64encode(buf.getvalue()).decode()
        icon_hash = hashlib.md5(icon_data.encode()).hexdigest()

        # 存储到数据库
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("INSERT OR IGNORE INTO icons (hash, data) VALUES (?, ?)", (icon_hash, icon_data))
            conn.commit()
            conn.close()
        except Exception:
            pass

        result = (icon_hash, icon_data)
        _icon_cache[process_name] = result
        return result
    except Exception:
        _icon_cache[process_name] = ("", "")
        return "", ""


def get_input_stats_snapshot() -> dict:
    """返回当前输入统计快照，存储到本地数据库，并返回需要上传的新 icon"""
    with _input_stats_lock:
        snapshot = {k: dict(v) for k, v in _app_stats.items()}
        _app_stats.clear()

    if not snapshot:
        return {}

    # 存储到本地数据库
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        timestamp = datetime.now().isoformat()

        for process_name, stats in snapshot.items():
            icon_hash, icon_data = extract_icon_base64(process_name)
            cursor.execute("""
                INSERT INTO input_stats (process, display_name, icon_hash, key_presses, left_clicks, right_clicks, scroll_distance, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                process_name,
                stats["display_name"],
                icon_hash,
                stats["key_presses"],
                stats["left_clicks"],
                stats["right_clicks"],
                stats["scroll_distance"],
                timestamp
            ))

        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Failed to save input stats: {e}")

    # 返回需要上传的新 icon（只返回本次快照中的 hash，Koishi 端会检查是否已存在）
    result = {}
    for process_name, stats in snapshot.items():
        icon_hash, icon_data = extract_icon_base64(process_name)
        result[process_name] = {
            "display_name": stats["display_name"],
            "icon_hash": icon_hash,
            "icon_base64": icon_data if icon_data else "",
            "key_presses": stats["key_presses"],
            "left_clicks": stats["left_clicks"],
            "right_clicks": stats["right_clicks"],
            "scroll_distance": stats["scroll_distance"],
        }
    return result


def save_activity_to_db(process: str, title: str):
    """保存活动记录到本地数据库"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO activity (process, title, timestamp)
            VALUES (?, ?, ?)
        """, (process, title, datetime.now().isoformat()))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Failed to save activity: {e}")


def query_daily_data(date_str: str = None) -> dict:
    """查询指定日期的活动和输入统计数据（供 Koishi 调用）"""
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")

    start_time = f"{date_str} 00:00:00"
    end_time = f"{date_str} 23:59:59"

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # 查询活动记录
        cursor.execute("""
            SELECT process, title, timestamp FROM activity
            WHERE timestamp BETWEEN ? AND ?
            ORDER BY timestamp
        """, (start_time, end_time))
        activities = [{"process": row[0], "title": row[1], "timestamp": row[2]} for row in cursor.fetchall()]

        # 查询输入统计
        cursor.execute("""
            SELECT process, display_name, icon_hash,
                   SUM(key_presses), SUM(left_clicks), SUM(right_clicks), SUM(scroll_distance)
            FROM input_stats
            WHERE timestamp BETWEEN ? AND ?
            GROUP BY process
        """, (start_time, end_time))
        input_stats = {}
        icon_hashes = set()
        for row in cursor.fetchall():
            input_stats[row[0]] = {
                "display_name": row[1],
                "icon_hash": row[2],
                "key_presses": row[3],
                "left_clicks": row[4],
                "right_clicks": row[5],
                "scroll_distance": row[6]
            }
            if row[2]:
                icon_hashes.add(row[2])

        # 查询浏览器活动
        cursor.execute("""
            SELECT domain, SUM(seconds) FROM browser_activity
            WHERE timestamp BETWEEN ? AND ?
            GROUP BY domain
        """, (start_time, end_time))
        browser_activity = {}
        for row in cursor.fetchall():
            browser_activity[row[0]] = row[1]

        # 查询图标数据
        icons = {}
        if icon_hashes:
            placeholders = ','.join('?' * len(icon_hashes))
            cursor.execute(f"SELECT hash, data FROM icons WHERE hash IN ({placeholders})", tuple(icon_hashes))
            for row in cursor.fetchall():
                icons[row[0]] = row[1]

        conn.close()
        return {
            "activities": activities,
            "input_stats": input_stats,
            "browser_activity": browser_activity,
            "icons": icons
        }
    except Exception as e:
        print(f"Failed to query daily data: {e}")
        return {"activities": [], "input_stats": {}, "browser_activity": {}, "icons": {}}


def get_browser_stats_snapshot() -> dict:
    """返回当前浏览器活动快照并存入本地数据库"""
    with _browser_stats_lock:
        snapshot = dict(_browser_stats)
        _browser_stats.clear()

    if not snapshot:
        return {}

    # 存储到本地数据库
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        timestamp = datetime.now().isoformat()
        for domain, seconds in snapshot.items():
            cursor.execute(
                "INSERT INTO browser_activity (domain, seconds, timestamp) VALUES (?, ?, ?)",
                (domain, seconds, timestamp)
            )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Failed to save browser stats: {e}")

    return snapshot


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
def _save_to_screen_dir(img):
    """按需截图保存到 screen 目录"""
    now = datetime.now()
    date_dir = Path(__file__).parent / "screen" / now.strftime("%Y-%m-%d")
    date_dir.mkdir(parents=True, exist_ok=True)
    filename = date_dir / f"screenshot_{now.strftime('%H%M%S')}.jpg"
    img.save(filename, "JPEG", quality=85)


def take_screenshot() -> str:
    img = pyautogui.screenshot()
    _save_to_screen_dir(img)
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
        img = pyautogui.screenshot()
    _save_to_screen_dir(img)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ── WebSocket Agent ───────────────────────────────────────────────────────────
class MonitorLunaAgent:
    def __init__(self):
        init_database()  # 初始化本地数据库
        self.config = load_config()
        self.status = "未连接"
        self.running = True
        self._loop = None
        self._ws = None
        self._last_window = None
        self._stats_task = None
        self._screenshot_task = None
        self._thread = None
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
            elif cmd == "query_daily_data":
                date_str = msg.get("date")
                data = await asyncio.get_event_loop().run_in_executor(None, query_daily_data, date_str)
                return {"type": "result", "id": cmd_id, "ok": True, "data": json.dumps(data, ensure_ascii=False)}
            else:
                return {"type": "result", "id": cmd_id, "ok": False, "error": f"unknown command: {cmd}"}
        except Exception as e:
            return {"type": "result", "id": cmd_id, "ok": False, "error": str(e)}

    async def _run_once(self) -> bool:
        """返回 True 表示曾成功连接过"""
        cfg = self.config
        url = cfg["url"]
        self.status = f"连接中... {url}"
        connected = False
        try:
            async with websockets.connect(url, open_timeout=10, ping_interval=30, ping_timeout=10) as ws:
                self._ws = ws
                await ws.send(json.dumps({"type": "hello", "token": cfg["token"], "device_id": cfg["device_id"]}))
                ack = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
                if ack.get("type") != "hello_ack":
                    self.status = f"握手失败: {ack.get('message', '未知错误')}"
                    return False
                connected = True
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
        return connected

    async def _window_monitor(self, ws, device_id: str):
        while self.running:
            try:
                info = await asyncio.get_event_loop().run_in_executor(None, get_window_info)
                key = (info["process"], info["title"])
                if key != self._last_window:
                    self._last_window = key
                    # 只存储到本地数据库
                    await asyncio.get_event_loop().run_in_executor(None, save_activity_to_db, info["process"], info["title"])
                if ws.closed:
                    break
            except asyncio.CancelledError:
                break
            except Exception:
                pass
            await asyncio.sleep(2)

    async def _input_stats_sender(self):
        """定期保存输入统计和浏览器活动到本地数据库"""
        try:
            while self.running:
                await asyncio.sleep(30)
                try:
                    await asyncio.get_event_loop().run_in_executor(None, get_input_stats_snapshot)
                except Exception:
                    pass
                try:
                    await asyncio.get_event_loop().run_in_executor(None, get_browser_stats_snapshot)
                except Exception:
                    pass
        except asyncio.CancelledError:
            pass

    async def _screenshot_task_loop(self):
        """本地定时截图任务"""
        try:
            while self.running:
                cfg = self.config
                if not cfg.get("screenshot_enabled", False):
                    await asyncio.sleep(60)
                    continue

                interval = cfg.get("screenshot_interval", 15) * 60
                try:
                    img = await asyncio.get_event_loop().run_in_executor(None, pyautogui.screenshot)
                    now = datetime.now()
                    date_dir = Path(__file__).parent / "cronscreen" / now.strftime("%Y-%m-%d")
                    date_dir.mkdir(parents=True, exist_ok=True)
                    filename = date_dir / f"screenshot_{now.strftime('%H%M')}.jpg"
                    await asyncio.get_event_loop().run_in_executor(None, lambda: img.save(filename, "JPEG", quality=85))
                except Exception as e:
                    print(f"Screenshot failed: {e}")

                await asyncio.sleep(interval)
        except asyncio.CancelledError:
            pass

    async def run_forever(self):
        delay = 3
        while self.running:
            was_connected = await self._run_once()
            if not self.running:
                break
            if was_connected:
                delay = 3
            self.status = f"重连中... ({delay}s 后)"
            await asyncio.sleep(delay)
            delay = min(delay * 2, 60)

    def start_in_thread(self):
        if self._thread and self._thread.is_alive():
            return

        def _thread():
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            # 启动独立的后台任务
            self._stats_task = self._loop.create_task(self._input_stats_sender())
            self._screenshot_task = self._loop.create_task(self._screenshot_task_loop())
            try:
                self._loop.run_until_complete(self.run_forever())
            except RuntimeError as e:
                if "Event loop stopped before Future completed" not in str(e):
                    raise
            finally:
                for task in (self._stats_task, self._screenshot_task):
                    if task:
                        task.cancel()
                pending = [task for task in asyncio.all_tasks(self._loop) if not task.done()]
                for task in pending:
                    task.cancel()
                if pending:
                    with contextlib.suppress(Exception):
                        self._loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
                self._loop.close()
                self._loop = None
                self._ws = None
                self._stats_task = None
                self._screenshot_task = None

        t = threading.Thread(target=_thread, daemon=True)
        self._thread = t
        t.start()

    def stop(self):
        self.running = False
        loop = self._loop
        if not loop:
            return

        if self._ws:
            asyncio.run_coroutine_threadsafe(self._ws.close(), loop)
        if self._stats_task:
            loop.call_soon_threadsafe(self._stats_task.cancel)
        if self._screenshot_task:
            loop.call_soon_threadsafe(self._screenshot_task.cancel)
        if self._thread and self._thread.is_alive() and threading.current_thread() is not self._thread:
            self._thread.join(timeout=3)


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
input[type="text"],input[type="password"],input[type="number"]{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}
input[type="checkbox"]{margin-right:8px}
.checkbox-label{display:flex;align-items:center;margin:16px 0}
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
<div class="checkbox-label">
<input type="checkbox" id="screenshot_enabled">
<label for="screenshot_enabled" style="margin:0">启用本地定时截图</label>
</div>
<label>截图间隔（分钟）</label>
<input type="number" id="screenshot_interval" min="1" max="120" value="15">
<label>浏览器扩展密码</label>
<input type="password" id="browser_token" placeholder="留空则不启用鉴权">
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
  document.getElementById('screenshot_enabled').checked=d.screenshot_enabled||false;
  document.getElementById('screenshot_interval').value=d.screenshot_interval||15;
  document.getElementById('browser_token').value=d.browser_token||'';
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
    device_id:document.getElementById('device_id').value,
    screenshot_enabled:document.getElementById('screenshot_enabled').checked,
    screenshot_interval:parseInt(document.getElementById('screenshot_interval').value),
    browser_token:document.getElementById('browser_token').value
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
    safe_config = dict(agent.config)
    safe_config["token"] = "***" if safe_config.get("token") else ""
    return web.json_response(safe_config)


async def handle_post_config(request):
    agent = request.app["agent"]
    data = await request.json()
    allowed_keys = {"url", "token", "device_id", "screenshot_enabled", "screenshot_interval", "browser_token"}
    filtered = {k: v for k, v in data.items() if k in allowed_keys}
    if not filtered.get("url", "").strip():
        return web.json_response({"ok": False, "message": "url is required"}, status=400)
    filtered["url"] = filtered.get("url", "").strip()
    filtered["device_id"] = filtered.get("device_id", "").strip() or agent.config.get("device_id", "my-pc")
    filtered["browser_token"] = filtered.get("browser_token", "").strip()
    if filtered.get("token") == "***":
        filtered["token"] = agent.config.get("token", "")
    if filtered.get("browser_token") == "***":
        filtered["browser_token"] = agent.config.get("browser_token", "")
    save_config({**agent.config, **filtered})
    agent.stop()
    agent.config = {**agent.config, **filtered}
    agent.running = True
    agent.start_in_thread()
    return web.json_response({"ok": True})


async def handle_status(request):
    agent = request.app["agent"]
    return web.json_response({"status": agent.status})


# ── 浏览器扩展 WebSocket 端点 ─────────────────────────────────────────────────
async def handle_browser_ws(request):
    """接收浏览器扩展的 WebSocket 连接，汇总活动数据"""
    agent = request.app["agent"]
    ws_response = web.WebSocketResponse()
    await ws_response.prepare(request)

    authenticated = False
    async for msg in ws_response:
        if msg.type == web.WSMsgType.TEXT:
            try:
                data = json.loads(msg.data)
            except Exception:
                continue

            if data.get("type") == "hello":
                browser_token = agent.config.get("browser_token", "")
                if browser_token and data.get("token") != browser_token:
                    await ws_response.send_json({"type": "error", "message": "invalid token"})
                    await ws_response.close()
                    break
                authenticated = True
                await ws_response.send_json({"type": "hello_ack", "message": "connected"})
                continue

            if not authenticated:
                await ws_response.send_json({"type": "error", "message": "send hello first"})
                continue

            if data.get("type") == "browser_activity":
                stats = data.get("stats", {})
                with _browser_stats_lock:
                    for domain, seconds in stats.items():
                        if isinstance(seconds, (int, float)) and seconds > 0:
                            _browser_stats[domain] = _browser_stats.get(domain, 0) + seconds
        elif msg.type in (web.WSMsgType.ERROR, web.WSMsgType.CLOSE):
            break

    return ws_response


def start_webui(agent: MonitorLunaAgent):
    app = web.Application()
    app["agent"] = agent
    app.router.add_get("/", handle_index)
    app.router.add_get("/api/config", handle_get_config)
    app.router.add_post("/api/config", handle_post_config)
    app.router.add_get("/api/status", handle_status)
    app.router.add_get("/ws/browser", handle_browser_ws)
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
