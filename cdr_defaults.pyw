# -*- coding: utf-8 -*-
"""
CorelDRAW 默认对象属性 · 悬浮控制窗

直接改写文档的「默认对象属性」样式（~graphic~ 图形 / ~artisticmedia~ 艺术笔），
此后用样条、贝塞尔、手绘、钢笔等工具绘制的新对象立即继承新样式。

已在 CorelDRAW Graphics Suite 2024 (25.0.0.230) / Python 3.12 实测通过。

几个来自实测的关键点（不要"想当然"改回去）：
  * cdrMillimeter 在 COM 层是 3，不是 VBA 文档写的 4；
  * StyleOutline.Width 的 getter 对小值恒返回 0，读取必须走 GetPropertyAsString("outline")；
  * Fill.UniformColor / PrimaryColor 均只读，填充色只能通过 SetPropertyAsString("fill", json) 写入；
  * CorelDRAW 不注册 ROT，GetActiveObject 不可用，只能用 Dispatch（会复用已运行实例）。
"""

import os
import re
import math
import glob
import shutil
import hashlib
import sys
import json
import time
import socket
import ctypes
import tempfile
import threading
import subprocess
import webbrowser
import urllib.parse
import urllib.request
import tkinter as tk
from tkinter import colorchooser, messagebox

try:
    import pythoncom
    import win32com.client as win32com
    # 固定用动态绑定：本机若存在 gen_py 类型库缓存，Dispatch 会切到早期绑定，
    # 参数校验与属性/方法判定随之改变，打包后（无缓存）行为就会不一致。
    from win32com.client import dynamic as win32dynamic
    HAS_COM = True
    COM_ERR = ""
except Exception as exc:                                   # pragma: no cover
    HAS_COM = False
    COM_ERR = str(exc)

try:
    import win32gui
    import win32api
    import win32con
    import win32process
    HAS_WIN32UI = True
except Exception:
    HAS_WIN32UI = False

try:
    import tray as traymod            # 纯 pywin32 托盘，不依赖 Pillow
    HAS_TRAY = True
except Exception as _tray_err:        # pragma: no cover
    traymod = None
    HAS_TRAY = False
    _TRAY_ERR = str(_tray_err)


def asset_path(name):
    """打包后在 _MEIPASS 里，源码模式在脚本目录。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

try:
    import psutil
    HAS_PSUTIL = True
except Exception:
    HAS_PSUTIL = False

try:
    import linecheck          # 线条闭合检测的纯几何部分
except Exception as _lc_err:  # pragma: no cover
    linecheck = None
    _LINECHECK_ERR = str(_lc_err)

try:
    import faces as facemod      # 平面图分面：按「面」找最窄通道
except Exception as _f_err:      # pragma: no cover
    facemod = None
    _FACES_ERR = str(_f_err)

try:
    import colors as colorname   # 颜色名映射卡
except Exception as _cn_err:     # pragma: no cover
    colorname = None
    _COLOR_ERR = str(_cn_err)

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
LOG_PATH = os.path.join(APP_DIR, "debug.log")

TENTH_PER_MM = 10000.0
NULL_GUID = "00000000-0000-0000-0000-000000000000"

STYLE_KEYS = {
    "graphic": "~graphic~",
    "artistic": "~artisticmedia~",
}

PRESET_WIDTHS = [("发丝", 0.0), ("0.1", 0.1), ("0.2", 0.2), ("0.35", 0.35),
                 ("0.5", 0.5), ("0.75", 0.75), ("1", 1.0)]

# 连不上 CorelDRAW 时的兜底色，连上后会被页面底部调色板替换
FALLBACK_COLORS = [
    (0, 0, 0), (64, 64, 64), (128, 128, 128), (192, 192, 192),
    (227, 30, 36), (245, 130, 32), (255, 199, 44), (0, 150, 64),
    (0, 110, 185), (140, 45, 140),
]

PAL_COLS = 14          # 每行色块数
PAL_CELL = 22          # 单格边长（含间距）
PAL_SW = 16            # 色块实际边长
PAL_ROWS = 3           # 可见行数
PAL_H = PAL_ROWS * PAL_CELL + 4

BG = "#F7F6F2"
FIELD_BG = "#FFFFFF"
LINE = "#E2E0D8"
CHIP_BG = "#EFEEE8"
CHIP_HOVER = "#E4E2DA"
TXT = "#242422"
MUTED = "#8A8880"
ACCENT = "#2F3E46"
ACCENT_HOVER = "#3C4E57"
ACCENT_SOFT = "#DDE4E6"
ACCENT_SOFT_HOVER = "#CFD9DC"
SEL_BG = "#3A4C54"
SEL_FG = "#FFFFFF"
TAG_BG = "#E9E7DF"
OK = "#2E7D5B"
WARN = "#C07A18"
BAD = "#B4423B"

FONT = ("Microsoft YaHei UI", 9)
FONT_S = ("Microsoft YaHei UI", 8)
FONT_B = ("Microsoft YaHei UI", 9, "bold")
FONT_NUM = ("Consolas", 10)

W, H = 336, 600
BAR_H = 36
QUICK_H = 32
VER_H = 26
DOCBAR_H = 28


def log(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------- 版本与更新
APP_VERSION = "1.1.3"

# 内网优先：先探 kust.work，通了就用内网源，不通才回落到公网 sjck.kust.work。
# 两个地址都能用环境变量覆盖，方便调试或换镜像。
UPDATE_PATH = "/cdr"
UPDATE_LAN_URL = os.environ.get("CDR_UPDATE_LAN") or ("https://kust.work" + UPDATE_PATH)
UPDATE_WAN_URL = os.environ.get("CDR_UPDATE_WAN") or ("https://sjck.kust.work" + UPDATE_PATH)
FEEDBACK_URL = os.environ.get("CDR_FEEDBACK_URL") or "https://sjck.kust.work/feedback"
UPDATE_TIMEOUT = 5.0
LAN_PROBE_TIMEOUT = 1.2


def version_tuple(text):
    """把 "v1.2.10" 这类字符串变成可比较的元组；读不到数字就返回 ()。"""
    nums = re.findall(r"\d+", str(text or ""))
    return tuple(int(n) for n in nums[:4]) if nums else ()


def url_reachable(url, timeout=LAN_PROBE_TIMEOUT):
    """探活：DNS + TCP 建连。比 ICMP ping 更贴近“这个地址能不能真拿到东西”。"""
    try:
        parts = urllib.parse.urlsplit(url)
        host = parts.hostname
        if not host:
            return False
        port = parts.port or (443 if parts.scheme == "https" else 80)
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def pick_update_source():
    """内网优先，返回 (url, 说明)。"""
    if url_reachable(UPDATE_LAN_URL):
        return UPDATE_LAN_URL, "内网 kust.work"
    return UPDATE_WAN_URL, "公网 sjck.kust.work"


def parse_update_payload(text):
    """尽量宽容地解析线上版本信息：JSON（version/latest/ver…）或纯文本 / HTML。

    支持的 JSON 形如：
        {"version": "1.3.0", "url_lan": "http://kust.work/cdr/CDR助手.exe",
         "url": "https://sjck.kust.work/cdr/CDR助手.exe", "notes": "…"}
    """
    info = {"version": "", "url": "", "url_lan": "", "notes": ""}
    raw = (text or "").strip()
    if not raw:
        return info
    try:
        data = json.loads(raw)
    except Exception:
        data = None
    if isinstance(data, dict):
        for key in ("version", "latest", "ver", "tag", "build"):
            v = data.get(key)
            if v:
                info["version"] = str(v).strip().lstrip("vV")
                break
        for key in ("url_lan", "lan", "url_intranet", "intranet"):
            if data.get(key):
                info["url_lan"] = str(data[key]).strip()
                break
        for key in ("url", "download", "link", "download_url"):
            if data.get(key):
                info["url"] = str(data[key]).strip()
                break
        for key in ("notes", "changelog", "desc", "description", "whatsnew"):
            if data.get(key):
                info["notes"] = str(data[key]).strip()
                break
    if not info["version"]:
        m = re.search(r"\bv?(\d+(?:\.\d+){1,3})\b", raw)
        if m:
            info["version"] = m.group(1)
    return info


def _fetch_from(url, label):
    req = urllib.request.Request(url, headers={
        "User-Agent": "CDRHelper/%s" % APP_VERSION,
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    })
    with urllib.request.urlopen(req, timeout=UPDATE_TIMEOUT) as resp:
        body = resp.read(256 * 1024)
    info = parse_update_payload(body.decode("utf-8", "ignore"))
    info["source"] = label
    info["source_url"] = url
    return info


def fetch_update_info():
    """先探内网，通了读内网；内网读失败或不通，回落公网。"""
    url, label = pick_update_source()
    if label.startswith("内网"):
        try:
            return _fetch_from(url, label)
        except Exception as exc:
            log("内网更新源失败(%r)，回落公网" % (exc,))
    return _fetch_from(UPDATE_WAN_URL, "公网 sjck.kust.work")


def pick_download_url(info):
    """下载地址：内网可达时优先内网的 exe 地址，其次公网，最后回到源地址目录。

    服务端可能只给相对路径（/cdr/cdr-helper.exe），也可能因为上游反代终止 TLS
    而把绝对地址写成 http —— 这里统一以「取版本时用的地址」为准纠正，
    避免 https 站点跳去 http 下载。
    """
    lan = str(info.get("source", "")).startswith("内网")
    base = info.get("source_url") or (UPDATE_LAN_URL if lan else UPDATE_WAN_URL)
    url = ""
    if lan and info.get("url_lan"):
        url = info["url_lan"]
    elif info.get("url"):
        url = info["url"]
    else:
        url = base.rstrip("/") + "/"

    if url and not url.lower().startswith(("http://", "https://")):
        url = urllib.parse.urljoin(base, url)
    if base.lower().startswith("https://") and url.lower().startswith("http://"):
        url = "https://" + url[len("http://"):]
    return url


def exe_path():
    """当前 exe 的绝对路径（源码模式返回 None）。"""
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return None


def cleanup_update_leftovers(only_old=False):
    """启动时清掉上次更新留下的 .old / .new。"""
    exe = exe_path()
    if not exe:
        return
    for suf in ((".old",) if only_old else (".old", ".new")):
        p = exe + suf
        try:
            if os.path.exists(p):
                os.remove(p)
                log("清理残留: %s" % p)
        except Exception as exc:
            log("清理 %s 失败: %r" % (p, exc))


def download_file(url, dest, on_progress=None, timeout=30):
    """下载到 dest，返回字节数。on_progress(got, total) 会在下载线程里被调用。"""
    req = urllib.request.Request(url, headers={
        "User-Agent": "CDRHelper/%s" % APP_VERSION,
        "Cache-Control": "no-cache",
    })
    tmp = dest + ".part"
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        got = 0
        with open(tmp, "wb") as f:
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if on_progress:
                    try:
                        on_progress(got, total)
                    except Exception:
                        pass
    if total and got != total:
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise IOError("下载不完整（%d/%d 字节）" % (got, total))
    os.replace(tmp, dest)
    return got


def looks_like_exe(path, min_size=512 * 1024):
    """粗校验：PE 头 + 体积。防的是"下载回来一个 404 页面"。"""
    try:
        if os.path.getsize(path) < min_size:
            return False
        with open(path, "rb") as f:
            return f.read(2) == b"MZ"
    except Exception:
        return False


def spawn_new_instance(path):
    """从 PyInstaller 单文件包里再启动一个自己。

    必须先清掉 _MEIPASS/_PYI_* 这些变量，否则新进程会沿用父进程的临时解包目录，
    父进程一退出目录被删，子进程就报 base_library.zip 找不到。
    """
    env = os.environ.copy()
    for k in list(env):
        if k in ("_MEIPASS", "_MEIPASS2", "_PYI_APPLICATION_HOME_DIR",
                 "_PYI_ARCHIVE_FILE", "_PYI_PARENT_PROCESS_LEVEL", "_PYI_APPLICATION_HOME_DIR"):
            env.pop(k, None)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return subprocess.Popen([path], env=env, close_fds=True)


def install_update(new_path):
    """把 new_path 换到当前 exe 上并重启。

    优先走「重命名运行中的 exe」这条路（同盘 rename 不受映像锁限制）；
    万一被杀软或权限拦住，退回到批处理：等本进程退出后再替换并启动。
    返回 True 表示已经交给替换流程，调用方应马上退出。
    """
    exe = exe_path()
    if not exe:
        return False

    old = exe + ".old"
    try:
        if os.path.exists(old):
            os.remove(old)
    except Exception:
        pass

    try:
        os.replace(exe, old)              # 运行中的 exe 允许改名
        os.replace(new_path, exe)         # 空位补上新文件
        spawn_new_instance(exe)
        log("更新完成：已改名旧版并用新版重启")
        return True
    except Exception as exc:
        log("原地替换失败(%r)，改用脚本替换" % (exc,))

    bat = os.path.join(tempfile.gettempdir(), "cdr_update_%d.bat" % os.getpid())
    lines = [
        "@echo off",
        "setlocal",
        ":retry",
        'move /y "%s" "%s" >nul 2>&1' % (new_path, exe),
        "if errorlevel 1 (",
        "  timeout /t 1 /nobreak >nul",
        "  goto retry",
        ")",
        'start "" "%s"' % exe,
        'del "%~f0"',
    ]
    try:
        with open(bat, "w", encoding="mbcs", errors="replace") as f:
            f.write("\r\n".join(lines) + "\r\n")
        flags = 0x00000008 | 0x08000000     # DETACHED_PROCESS | CREATE_NO_WINDOW
        subprocess.Popen(["cmd", "/c", bat], creationflags=flags, close_fds=True)
        log("更新完成：已交给批处理替换")
        return True
    except Exception as exc:
        log("脚本替换也失败: %r" % (exc,))
        return False


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "CorelDrawHelper"


# ---------------------------------------------------------------- 单实例
MUTEX_NAME = "Local\\CorelDrawHelper_SingleInstance"
ACTIVATE_FLAG = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
                             "CDR助手", "activate.flag")


def acquire_single_instance():
    """拿到互斥体返回句柄；已有实例在跑返回 None；出错返回 -1（放行）。"""
    try:
        k32 = ctypes.windll.kernel32
        k32.CreateMutexW.restype = ctypes.c_void_p
        h = k32.CreateMutexW(None, False, MUTEX_NAME)
        if int(k32.GetLastError()) == 183:      # ERROR_ALREADY_EXISTS
            return None
        return h
    except Exception:
        return -1


def cdr_pids():
    """活着的 CorelDRAW 进程（排除僵尸/已退出未回收的）。"""
    if not HAS_PSUTIL:
        return []
    out = []
    for p in psutil.process_iter(["pid", "name", "status"]):
        if "CorelDRW" not in (p.info.get("name") or ""):
            continue
        if p.info.get("status") == psutil.STATUS_ZOMBIE:
            continue
        try:
            if not p.is_running():
                continue
        except Exception:
            continue
        out.append(p.info["pid"])
    return out


def has_windowed_cdr():
    """有没有一个带可见窗口的 CorelDRAW（= 用户真正在用的那个）。"""
    visible = set()
    try:
        def cb(h, _):
            try:
                if win32gui.IsWindowVisible(h):
                    visible.add(win32process.GetWindowThreadProcessId(h)[1])
            except Exception:
                pass
        win32gui.EnumWindows(cb, None)
    except Exception:
        return bool(cdr_pids())
    return bool(visible & set(cdr_pids()))


def headless_cdr_pids(min_age=20):
    """只认「命令行带 /Auto 且没有可见窗口」的自动化实例。

    用户自己开出来的 CorelDRAW 一律不碰——宁可让他手动关，也不能误杀正在编辑的文件。
    """
    if not HAS_PSUTIL:
        return []
    visible = set()
    try:
        def cb(h, _):
            try:
                if win32gui.IsWindowVisible(h):
                    visible.add(win32process.GetWindowThreadProcessId(h)[1])
            except Exception:
                pass
        win32gui.EnumWindows(cb, None)
    except Exception:
        return []
    now = time.time()
    out = []
    for p in psutil.process_iter(["pid", "name", "create_time", "cmdline"]):
        nm = p.info.get("name") or ""
        if "CorelDRW" not in nm:
            continue
        if p.info["pid"] in visible:
            continue
        cmd = " ".join(p.info.get("cmdline") or []).lower()
        if "/auto" not in cmd:
            continue
        try:
            if now - float(p.info["create_time"]) < min_age:
                continue
        except Exception:
            pass
        out.append(p.info["pid"])
    return out


def kill_pids(pids):
    n = 0
    for pid in pids:
        try:
            p = psutil.Process(pid)
            p.kill()
            n += 1
            log("结束无窗口的 CorelDRAW 残留实例 pid=%s" % pid)
        except Exception:
            pass
    return n


def find_app_window():
    found = []

    def cb(h, _):
        try:
            if win32gui.GetClassName(h) == "TkTopLevel":
                t = win32gui.GetWindowText(h)
                if "CorelDRAW 助手" in t:
                    found.append(h)
        except Exception:
            pass
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        return 0
    return found[0] if found else 0


def center_window(hwnd):
    try:
        r = win32gui.GetWindowRect(hwnd)
        w, h = r[2] - r[0], r[3] - r[1]
        sw, sh = win32gui.GetSystemMetrics(0), win32gui.GetSystemMetrics(1)
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        win32gui.ShowWindow(hwnd, 9)                 # SW_RESTORE
        win32gui.SetWindowPos(hwnd, -1, x, y, 0, 0, 0x0001)   # HWND_TOPMOST + NOSIZE
        win32gui.SetForegroundWindow(hwnd)
        try:
            win32gui.FlashWindow(hwnd, True)
        except Exception:
            pass
        return True
    except Exception:
        return False


def notify_already_running():
    """第二个实例：把已有窗口拉到屏幕中央，提示一句，然后自己退出。"""
    hwnd = find_app_window()
    if hwnd:
        center_window(hwnd)
    try:
        os.makedirs(os.path.dirname(ACTIVATE_FLAG), exist_ok=True)
        with open(ACTIVATE_FLAG, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    except Exception:
        pass

    top = tk.Tk()
    top.overrideredirect(True)
    top.attributes("-topmost", True)
    top.configure(bg=TAG_BG)
    tk.Label(top, text="CorelDRAW 助手已在运行\n已把窗口放到屏幕中央",
             bg=TAG_BG, fg=TXT, font=FONT_S, padx=18, pady=12).pack()
    top.update_idletasks()
    w, h = top.winfo_reqwidth(), top.winfo_reqheight()
    sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
    top.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, (sh - h) // 2 - 120))
    top.after(2200, top.destroy)
    top.mainloop()


# ---------------------------------------------------------------- 批处理备份
# 每次批处理前整份复制一份 .cdr，每个文档留最近 10 份；
# 文档「保存并关闭」后自动清空该文档的备份缓存。
BACKUP_ROOT = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
                           "CDR助手", "backup")
BACKUP_KEEP = 10


def _backup_dir(path):
    key = hashlib.md5(os.path.normcase(os.path.abspath(path)).encode("utf-8")).hexdigest()[:12]
    return os.path.join(BACKUP_ROOT, key)


def backup_file(path, keep=BACKUP_KEEP):
    """整份复制；超过 keep 份时删最旧的。返回备份路径或 None。"""
    if not path or not os.path.isfile(path):
        return None
    d = _backup_dir(path)
    try:
        os.makedirs(d, exist_ok=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        dest = os.path.join(d, "%s_%s.cdr" % (stem, time.strftime("%Y%m%d_%H%M%S")))
        if os.path.exists(dest):
            dest = os.path.join(d, "%s_%s_%03d.cdr"
                                % (stem, time.strftime("%Y%m%d_%H%M%S"),
                                   int(time.time() * 1000) % 1000))
        shutil.copy2(path, dest)
        files = sorted(glob.glob(os.path.join(d, "*.cdr")), key=os.path.getmtime)
        for old in files[:-keep]:
            try:
                os.remove(old)
            except Exception:
                pass
        return dest
    except Exception as exc:
        log("备份失败 %r" % (exc,))
        return None


def list_backups(path):
    d = _backup_dir(path)
    if not os.path.isdir(d):
        return []
    return sorted(glob.glob(os.path.join(d, "*.cdr")), key=os.path.getmtime)


def clear_backups(path):
    try:
        shutil.rmtree(_backup_dir(path), ignore_errors=True)
    except Exception:
        pass


def launch_command():
    """返回写进注册表 Run 键的启动命令行。"""
    if getattr(sys, "frozen", False):
        return '"%s"' % os.path.abspath(sys.executable)
    exe = sys.executable
    if exe.lower().endswith("python.exe"):
        exe = exe[:-len("python.exe")] + "pythonw.exe"
    return '"%s" "%s"' % (exe, os.path.abspath(__file__))


def autostart_enabled():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            v, _ = winreg.QueryValueEx(k, RUN_NAME)
            return bool(str(v).strip())
    except Exception:
        return False


def set_autostart(enable):
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                            winreg.KEY_SET_VALUE) as k:
            if enable:
                winreg.SetValueEx(k, RUN_NAME, 0, winreg.REG_SZ, launch_command())
            else:
                try:
                    winreg.DeleteValue(k, RUN_NAME)
                except FileNotFoundError:
                    pass
        return True, ""
    except Exception as exc:
        return False, "写入启动项失败：%s" % exc


def rgb_hex(rgb):
    return "#%02X%02X%02X" % tuple(rgb)


def hex_to_rgb(h):
    try:
        h = str(h).strip().lstrip("#")
        if len(h) >= 6:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        pass
    return None


def parse_color_token(token):
    """解析 Corel 颜色串：RGB255,USER,R,G,B,100,GUID / CMYK,USER,C,M,Y,K,100,GUID"""
    try:
        parts = str(token).split(",")
        model = parts[0].upper()
        if model == "RGB255":
            return (int(parts[2]), int(parts[3]), int(parts[4]))
        if model == "CMYK":
            c, m, y, k = (int(parts[i]) for i in (2, 3, 4, 5))
            return cmyk_to_rgb(c, m, y, k)
        if model in ("GRAY", "GRAY255"):
            v = int(parts[2])
            if v <= 100:
                v = int(round(v * 255 / 100.0))
            return (v, v, v)
    except Exception:
        pass
    return None


def cmyk_to_rgb(c, m, y, k):
    c, m, y, k = [max(0, min(100, v)) / 100.0 for v in (c, m, y, k)]
    r = 255 * (1 - c) * (1 - k)
    g = 255 * (1 - m) * (1 - k)
    b = 255 * (1 - y) * (1 - k)
    return (int(round(r)), int(round(g)), int(round(b)))


def rgb_to_cmyk(rgb):
    r, g, b = [v / 255.0 for v in rgb]
    k = 1.0 - max(r, g, b)
    if k >= 0.999:
        return (0, 0, 0, 100)
    c = (1 - r - k) / (1 - k)
    m = (1 - g - k) / (1 - k)
    y = (1 - b - k) / (1 - k)
    return tuple(int(round(v * 100)) for v in (c, m, y, k))


# ---------------------------------------------------------------- COM 桥

class CdrBridge(object):

    def __init__(self):
        self.app = None
        self.doc = None
        self.connected = False
        self.doc_name = ""
        self.mm_unit = 3
        self._unit_ready = False
        self._line_geo = None      # 线条几何缓存（按当前文档）

    # ---- 进程探测 ---------------------------------------------

    @staticmethod
    def cdr_running():
        if not HAS_PSUTIL:
            return None
        try:
            for p in psutil.process_iter(["name"]):
                if (p.info.get("name") or "").lower() == "coreldrw.exe":
                    return True
        except Exception:
            return None
        return False

    # ---- 连接 -------------------------------------------------

    def attach(self):
        if not HAS_COM:
            return False, "缺少 pywin32，无法调用 CorelDRAW"

        running = self.cdr_running()
        if running is False and not self.connected:
            return False, "CorelDRAW 未运行，请先打开"

        # 只有「带可见窗口」的实例才是用户真正在用的。
        # 没有窗口时：清掉 /Auto 自动化残留，但绝不碰用户自己开的实例。
        if not has_windowed_cdr():
            auto = headless_cdr_pids()
            killed = kill_pids(auto) if auto else 0
            if killed:
                time.sleep(0.4)
            if not has_windowed_cdr():
                self.connected = False
                return False, ("CorelDRAW 窗口还没就绪%s，请稍后重试"
                               % (("（已清理 %d 个自动化残留）" % killed) if killed else ""))

        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            self.app = win32dynamic.Dispatch("CorelDRAW.Application")
        except Exception as exc:
            log("Dispatch 失败: %r" % (exc,))
            self.connected = False
            return False, "连接 CorelDRAW 失败"

        try:
            n_docs = int(self.app.Documents.Count)
        except Exception:
            n_docs = -1
        if n_docs <= 0:
            self.doc = None
            self.connected = True
            self.doc_name = ""
            return True, "已连接 · 等待文档"
        try:
            self.doc = self.app.ActiveDocument
            self.doc_name = str(self.doc.Name)
            self.connected = True
            self._unit_ready = False
            return True, "已连接 · " + self.doc_name[:18]
        except Exception:
            self.connected = False
            return False, "读取活动文档失败"

    def ensure_doc(self):
        """每次操作前刷新活动文档，支持用户切换文档。"""
        if self.app is None:
            return False
        try:
            if int(self.app.Documents.Count) == 0:
                self.doc = None
                return False
            self.doc = self.app.ActiveDocument
            self.doc_name = str(self.doc.Name)
            return True
        except Exception:
            self.doc = None
            return False

    # ---- 单位 -------------------------------------------------

    def resolve_mm_unit(self):
        """确定 cdrMillimeter 在 COM 层的实际枚举值。"""
        if self._unit_ready:
            return self.mm_unit
        doc = self.doc
        if doc is None:
            return 3

        try:
            w = float(doc.Pages.Item(1).SizeWidth)
        except Exception:
            w = 0.0

        if w > 0:
            for u in range(1, 17):
                try:
                    mm = float(doc.FromUnits(w, u))
                except Exception:
                    continue
                if 100.0 <= mm <= 420.0:
                    self.mm_unit = u
                    self._unit_ready = True
                    return u

        for u in range(1, 17):
            try:
                if abs(float(doc.ToUnits(25.4, u)) - 1.0) < 1e-6:
                    self.mm_unit = u
                    self._unit_ready = True
                    return u
                if abs(float(doc.ToUnits(1.0, u)) - 1.0) < 1e-6:
                    self.mm_unit = u
                    self._unit_ready = True
                    return u
            except Exception:
                continue

        self.mm_unit = 3
        self._unit_ready = True
        return 3

    def to_doc(self, mm):
        try:
            return float(self.doc.ToUnits(float(mm), self.mm_unit))
        except Exception:
            return float(mm)

    def from_doc(self, value):
        try:
            return float(self.doc.FromUnits(float(value), self.mm_unit))
        except Exception:
            return float(value)

    # ---- 样式 -------------------------------------------------

    def find_style(self, key):
        if self.doc is None:
            return None
        target = STYLE_KEYS.get(key)
        if not target:
            return None
        sheet = None
        try:
            sheet = self.doc.StyleSheet
        except Exception:
            return None

        try:
            st = sheet.FindStyle(target)
            if st is not None:
                return st
        except Exception as exc:
            log("FindStyle(%s) -> %r" % (target, exc))

        try:
            od = sheet.ObjectDefaults
            for i in range(1, int(od.Count) + 1):
                try:
                    item = od.Item(i)
                    if str(item.Name) == target:
                        return item
                except Exception:
                    continue
        except Exception as exc:
            log("遍历 ObjectDefaults 失败: %r" % (exc,))
        return None

    # ---- 读取 -------------------------------------------------

    def _outline_json(self, st):
        return json.loads(str(st.GetPropertyAsString("outline")))

    def _fill_json(self, st):
        return json.loads(str(st.GetPropertyAsString("fill")))

    def _dump_palette(self, pal, limit=400):
        out = []
        try:
            total = int(pal.ColorCount)
        except Exception:
            return out
        for i in range(1, total + 1):
            if len(out) >= limit:
                break
            try:
                rgb = hex_to_rgb(pal.Color(i).HexValue)
            except Exception:
                continue
            if rgb:
                out.append(rgb)
        return out

    def read_palette(self):
        """读取页面底部调色板：优先文档调色板，其次标记为 Default 的那个。"""
        if self.doc is None:
            return []

        colors = []
        try:
            dp = self.doc.Palette
            if dp is not None and int(dp.ColorCount) >= 8:
                colors = self._dump_palette(dp)
        except Exception as exc:
            log("读文档调色板失败: %r" % (exc,))

        if not colors and self.app is not None:
            try:
                pals = self.app.Palettes
                for i in range(1, int(pals.Count) + 1):
                    try:
                        p = pals.Item(i)
                        if bool(p.Default) and int(p.ColorCount) > 0:
                            colors = self._dump_palette(p)
                            if colors:
                                break
                    except Exception:
                        continue
            except Exception as exc:
                log("读默认调色板失败: %r" % (exc,))

        return colors[:400]

    def list_docs(self):
        """列出所有打开的文档：[(index, name)]"""
        out = []
        if self.app is None:
            return out
        try:
            n = int(self.app.Documents.Count)
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                out.append((i, str(self.app.Documents.Item(i).Name)))
            except Exception:
                continue
        return out

    def use_doc(self, index, activate=False):
        """把内部指针切到指定文档。"""
        if self.app is None:
            return False
        try:
            d = self.app.Documents.Item(index)
        except Exception:
            return False
        if activate:
            try:
                d.Activate()
            except Exception:
                pass
        self.doc = d
        try:
            self.doc_name = str(d.Name)
        except Exception:
            self.doc_name = ""
        self._unit_ready = False
        self.connected = True
        return True

    def read_current(self):
        st = self.find_style("graphic")
        if st is None:
            return None

        data = {}
        try:
            w = int(self._outline_json(st).get("width", 0)) / TENTH_PER_MM
            data["width_mm"] = 0.0 if w < 0.005 else round(w, 4)
        except Exception as exc:
            log("读线宽失败: %r" % (exc,))
            data["width_mm"] = None

        try:
            data["outline_rgb"] = parse_color_token(self._outline_json(st).get("color"))
        except Exception:
            data["outline_rgb"] = None

        o = st.Outline
        for key, attr, cast in (
            ("line_caps", "LineCaps", int),
            ("line_join", "LineJoin", int),
            ("justify", "Justification", int),
            ("scale_with", "ScaleWithShape", bool),
            ("behind_fill", "BehindFill", bool),
        ):
            try:
                data[key] = cast(getattr(o, attr))
            except Exception:
                data[key] = None

        try:
            fj = self._fill_json(st)
            if int(fj.get("type", 0)) == 0:
                data["fill_mode"] = "none"
                data["fill_rgb"] = None
            else:
                data["fill_mode"] = "custom"
                data["fill_rgb"] = parse_color_token(fj.get("primaryColor"))
        except Exception:
            data["fill_mode"] = None
            data["fill_rgb"] = None
        return data

    # ---- 写入 -------------------------------------------------

    def apply(self, state):
        if not self.connected or self.app is None:
            return 0, "未连接 CorelDRAW"
        if not self.ensure_doc():
            return 0, "没有打开的文档"

        kinds = [k for k, on in state.get("kinds", {}).items() if on]
        if not kinds:
            return 0, "请至少勾选一个对象类型"

        self.resolve_mm_unit()
        done, missing = 0, []
        for kind in kinds:
            st = self.find_style(kind)
            if st is None:
                missing.append(kind)
                continue
            try:
                self._apply_one(st, state)
                done += 1
            except Exception as exc:
                log("apply %s 失败: %r" % (kind, exc))
                missing.append(kind)

        if done and missing:
            return done, "已写入 %d 项，%d 项未找到" % (done, len(missing))
        if done:
            return done, "已写入默认样式 · " + self.doc_name[:16]
        return 0, "未找到默认样式"

    def _apply_one(self, st, state):
        mm = state.get("width_mm")
        if mm is not None:
            if float(mm) <= 0:
                st.Outline.Width = self.to_doc(0.0001)      # 发丝线
            else:
                st.Outline.Width = self.to_doc(mm)

        rgb = state.get("outline_rgb")
        if rgb:
            try:
                st.Outline.Type = 1
            except Exception:
                pass
            st.Outline.Color = self.app.CreateRGBColor(*rgb)

        mode = state.get("fill_mode")
        if mode == "none":
            try:
                st.Fill.Type = 0
            except Exception:
                pass
        elif mode in ("outline", "custom"):
            color = rgb if mode == "outline" else state.get("fill_rgb")
            if color:
                self._write_fill(st, tuple(color))

        o = st.Outline
        for key, attr, cast in (
            ("line_caps", "LineCaps", int),
            ("line_join", "LineJoin", int),
            ("justify", "Justification", int),
            ("scale_with", "ScaleWithShape", bool),
            ("behind_fill", "BehindFill", bool),
        ):
            val = state.get(key)
            if val is None:
                continue
            try:
                setattr(o, attr, cast(val))
            except Exception as exc:
                log("写 %s 失败: %r" % (attr, exc))

    def apply_linestyle(self, index):
        """线型只能作用于对象（默认对象属性里没有线型字段）。"""
        if not self.ensure_doc():
            return 0, "没有打开的文档"
        try:
            sel = self.doc.SelectionRange
            count = int(sel.Count)
        except Exception as exc:
            log("读选区失败: %r" % (exc,))
            return 0, "无法读取选区"
        if count == 0:
            return 0, "请先在 CorelDRAW 里选中图形"

        styles = self.app.OutlineStyles
        done = 0
        for i in range(1, count + 1):
            try:
                sel.Item(i).Outline.Style = styles.Item(index)
                done += 1
            except Exception as exc:
                log("线型应用到第 %d 个对象失败: %r" % (i, exc))
        if done == 0:
            return 0, "选中的对象不支持改线型"
        return done, "已改 %d 个对象的线型" % done

    # ---- 批量：按线型分组 -------------------------------------

    @staticmethod
    def _dash_count(shape):
        """0 = 实线，>0 = 虚线；返回 None 表示不参与分组（无轮廓 / 已经是组合）。

        线型判据取 Outline.Style.DashCount：CorelDRAW 25 的 LineStyle 上没有
        Name，Application.LineStyles 也不存在（集合挂在 Application.OutlineStyles，
        且只能按 Index 取），DashCount 是唯一可靠的读数。
        """
        try:
            if int(shape.Type) == 7:          # 已是组合，不再往组合里套组合
                return None
            o = shape.Outline
            if int(o.Type) == 0:              # 无轮廓：文本、位图、无边框对象
                return None
            return int(o.Style.DashCount)
        except Exception:
            return None

    def group_by_linetype(self, progress=None):
        """把每个图层里的线条按「实线 / 虚线」各合成一个组合。

        两个必须守住的点（都是实测踩出来的）：
          * 分组前必须 layer.Activate()。CreateShapeRangeFromArray().Group() 会把
            新组合并到「当前活动图层」，不激活就等于把别的图层的对象搬走；
          * 每处理一类都要重新读 layer.Shapes.Count，分组后旧引用会失效。
        """
        if not self.connected or self.app is None:
            return None, "未连接 CorelDRAW"
        if not self.ensure_doc():
            return None, "没有打开的文档"

        doc = self.doc
        try:
            page_count = int(doc.Pages.Count)
        except Exception:
            page_count = 1

        total_units = 0
        for pno in range(1, page_count + 1):
            try:
                total_units += int(doc.Pages.Item(pno).Layers.Count)
            except Exception:
                pass
        total_units = max(1, total_units)
        done_units = 0

        rows = []
        try:
            doc.BeginCommandGroup("按线型分为实线/虚线两组")
        except Exception:
            pass
        try:
            for pno in range(1, page_count + 1):
                try:
                    page = doc.Pages.Item(pno)
                    layer_count = int(page.Layers.Count)
                except Exception as exc:
                    log("读第 %d 页失败: %r" % (pno, exc))
                    continue
                try:
                    page.Activate()        # 多页文档：先切到该页，否则组会落到别的页
                except Exception:
                    pass

                for lno in range(1, layer_count + 1):
                    done_units += 1
                    if progress:
                        progress(done_units / float(total_units),
                                 "第 %d 页 · 图层 %d/%d" % (pno, lno, layer_count))
                    try:
                        layer = page.Layers.Item(lno)
                        name = str(layer.Name)
                        total = int(layer.Shapes.Count)
                    except Exception:
                        continue
                    if total == 0:
                        continue
                    try:
                        if not bool(layer.Editable):
                            rows.append((pno, name, 0, 0, "图层锁定"))
                            continue
                    except Exception:
                        pass
                    try:
                        layer.Activate()               # 关键：先激活目标图层
                    except Exception as exc:
                        log("激活图层 %s 失败: %r" % (name, exc))
                        rows.append((pno, name, 0, 0, "无法激活"))
                        continue

                    n_solid = n_dashed = 0
                    note = ""
                    for want_dash, gname in ((False, "实线"), (True, "虚线")):
                        targets = []
                        try:
                            for i in range(1, int(layer.Shapes.Count) + 1):
                                sh = layer.Shapes.Item(i)
                                dc = self._dash_count(sh)
                                if dc is None:
                                    continue
                                if (dc > 0) == want_dash:
                                    targets.append(sh)
                        except Exception as exc:
                            log("扫描图层 %s 失败: %r" % (name, exc))
                            note = "扫描失败"
                            break

                        if want_dash:
                            n_dashed = len(targets)
                        else:
                            n_solid = len(targets)
                        if len(targets) < 2:           # 单对象不建组
                            continue
                        try:
                            grp = doc.CreateShapeRangeFromArray(targets).Group()
                            try:
                                grp.Name = gname
                            except Exception:
                                pass
                        except Exception as exc:
                            log("图层 %s 的%s分组失败: %r" % (name, gname, exc))
                            note = "部分分组失败"
                    rows.append((pno, name, n_solid, n_dashed, note))
        except Cancelled:
            raise
        except Exception as exc:
            log("group_by_linetype 异常: %r" % (exc,))
            return None, "分组失败：%s" % exc
        finally:
            try:
                doc.EndCommandGroup()
            except Exception:
                pass

        if not rows:
            return None, "这个文档没有可处理的线条"

        ok_rows = [r for r in rows if (r[2] >= 2 or r[3] >= 2) and not r[4]]
        if not ok_rows:
            return None, "没有需要分组的图层（每类都不足两个对象）"
        s = sum(r[2] for r in ok_rows if r[2] >= 2)
        d = sum(r[3] for r in ok_rows if r[3] >= 2)
        return rows, "已分组 %d 个图层 · 实线 %d 条 / 虚线 %d 条" % (len(ok_rows), s, d)

    # ---- 线条闭合检测（边界 / 全图） ---------------------------

    LINE_SKIP_LAYERS = ("卫星图", "底图", "检测标记")
    BOUNDARY_LAYER = "边界"
    MARK_LAYER = "检测标记"

    def _mmu(self):
        return self.resolve_mm_unit()

    def doc_to_mm(self, v):
        """文档单位 → 毫米。注意 CorelDRAW 的 FromUnits/ToUnits 方向与名字相反：
        ToUnits(值, 单位) 把「以该单位表示的值」换成文档单位；
        FromUnits(值, 单位) 把「文档单位的值」换成该单位。"""
        try:
            return float(self.doc.FromUnits(float(v), self._mmu()))
        except Exception:
            return float(v) * 25.4

    def mm_to_doc(self, v):
        try:
            return float(self.doc.ToUnits(float(v), self._mmu()))
        except Exception:
            return float(v) / 25.4

    @staticmethod
    def _bbox(sh):
        try:
            return tuple(float(v) for v in sh.GetBoundingBox(False, False))
        except Exception:
            return None

    @staticmethod
    def _match_addr(cands, bb, tol=1e-4):
        """按 bbox 把临时文档里的对象对回源对象（粘贴顺序不保证一致）。"""
        if bb is None:
            return None
        best = None
        for item in cands:
            if item[2] or item[1] is None:
                continue
            d = max(abs(a - b) for a, b in zip(bb, item[1]))
            if d <= tol and (best is None or d < best[0]):
                best = (d, item)
        if best is None:
            return None
        best[1][2] = True
        return best[1][0]

    @staticmethod
    def _walk(shapes, prefix=()):
        out = []
        try:
            n = int(shapes.Count)
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                sh = shapes.Item(i)
            except Exception:
                continue
            addr = prefix + (i,)
            if int(sh.Type) == 7:
                out.extend(CdrBridge._walk(sh.Shapes, addr))
            else:
                out.append((addr, sh))
        return out

    @staticmethod
    def _read_polyline(shape):
        """返回 [(closed, [(x, y), ...]), ...]（文档单位）。

        优先用 SubPath.GetCurveInfo()：一次调用拿到全部节点坐标；
        拿不到再退回逐节点读取。
        """
        out = []
        try:
            c = shape.Curve
            n = int(c.SubPaths.Count)
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                sp = c.SubPaths.Item(i)
                closed = bool(sp.Closed)
                pts = []
                try:
                    info = sp.GetCurveInfo()
                    for nd in info:
                        pts.append((float(nd.PositionX), float(nd.PositionY)))
                except Exception:
                    nodes = sp.Nodes
                    for k in range(1, int(nodes.Count) + 1):
                        nd = nodes.Item(k)
                        pts.append((float(nd.PositionX), float(nd.PositionY)))
                if len(pts) >= 2:
                    out.append((closed, pts))
            except Exception:
                continue
        return out

    @staticmethod
    def _bbox(sh):
        try:
            return tuple(float(v) for v in sh.GetBoundingBox(False, False))
        except Exception:
            return None

    @staticmethod
    def _match_addr(cands, bb, tol=1e-4):
        """按 bbox 把临时文档里的对象对回源对象（粘贴顺序不保证一致）。"""
        if bb is None:
            return None
        best = None
        for item in cands:
            if item[2] or item[1] is None:
                continue
            d = max(abs(a - b) for a, b in zip(bb, item[1]))
            if d <= tol and (best is None or d < best[0]):
                best = (d, item)
        if best is None:
            return None
        best[1][2] = True
        return best[1][0]

    @staticmethod
    def _walk(shapes, prefix=()):
        out = []
        try:
            n = int(shapes.Count)
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                sh = shapes.Item(i)
            except Exception:
                continue
            addr = prefix + (i,)
            if int(sh.Type) == 7:
                out.extend(CdrBridge._walk(sh.Shapes, addr))
            else:
                out.append((addr, sh))
        return out

    @staticmethod
    def _read_polyline(shape):
        """返回 [(closed, [(x, y), ...]), ...]（文档单位）"""
        out = []
        try:
            c = shape.Curve
            n = int(c.SubPaths.Count)
        except Exception:
            return out
        for i in range(1, n + 1):
            try:
                sp = c.SubPaths.Item(i)
            except Exception:
                continue
            pts = []
            try:
                raw = sp.GetPolyline(0)
                if raw and isinstance(raw[0], (list, tuple)):
                    pts = [(float(p[0]), float(p[1])) for p in raw]
                else:
                    for k in range(0, len(raw) - 1, 2):
                        pts.append((float(raw[k]), float(raw[k + 1])))
            except Exception:
                try:
                    nodes = sp.Nodes
                    for k in range(1, int(nodes.Count) + 1):
                        nd = nodes.Item(k)
                        pts.append((float(nd.PositionX), float(nd.PositionY)))
                except Exception:
                    pts = []
            if len(pts) >= 2:
                out.append((bool(sp.Closed), pts))
        return out

    def _ensure_curve(self, shapes, addr):
        """按地址取对象；不是曲线就先转曲。返回曲线对象或 None。"""
        obj = shapes
        for k in addr[:-1]:
            try:
                obj = obj.Item(k).Shapes
            except Exception:
                return None
        try:
            sh = obj.Item(addr[-1])
        except Exception:
            return None
        if int(sh.Type) != 3:
            try:
                sh.ConvertToCurves()
            except Exception:
                return None
            try:
                sh = obj.Item(addr[-1])
            except Exception:
                return None
        return sh if int(sh.Type) == 3 else None

    def build_line_geo(self, progress=None):
        """在当前文档里临时转曲 → 读折线（毫米）→ 一次撤销还原。

        不再复制到临时文档：副本粘贴会丢对象、也可能把坐标弄歪，
        原地转曲 + Undo 能保证读到的就是图上真实位置。
        progress(frac, text)；抛 Cancelled 可中断（同样会撤销）。
        """
        if linecheck is None:
            return {}
        doc = self.doc
        page = doc.ActivePage
        todo = []
        for li in range(1, int(page.Layers.Count) + 1):
            L = page.Layers.Item(li)
            nm = str(L.Name)
            if nm in self.LINE_SKIP_LAYERS or int(L.Shapes.Count) == 0:
                continue
            todo.append((nm, L))
        n_layer = max(1, len(todo))
        geo = {}
        grouped = False
        try:
            doc.BeginCommandGroup("线条几何采集(临时)")
            grouped = True
        except Exception:
            pass
        try:
            for li, (name, L) in enumerate(todo):
                leaves = self._walk(L.Shapes)      # 地址先取，转曲不改索引
                paths = []
                for k, (addr, sh) in enumerate(leaves):
                    if progress:
                        progress((li + k / float(max(1, len(leaves)))) / n_layer,
                                 "读取 %s  %d/%d" % (name, k + 1, len(leaves)))
                    try:
                        if int(sh.Type) != 3:
                            sh.ConvertToCurves()
                        s2 = self._get_by_addr(L.Shapes, addr)
                        if s2 is None:
                            continue
                        for sub_i, (closed, pts) in enumerate(self._read_polyline(s2)):
                            paths.append({
                                "addr": list(addr),
                                "sub": sub_i,
                                "closed": closed,
                                "pts": [(self.doc_to_mm(x), self.doc_to_mm(y))
                                        for x, y in pts],
                            })
                    except Cancelled:
                        raise
                    except Exception as exc:
                        log("采集 %s 第 %d 个失败: %r" % (name, k, exc))
                geo[name] = paths
        finally:
            try:
                doc.EndCommandGroup()
            except Exception:
                pass
            if grouped:
                try:
                    doc.Undo(1)                # 一次撤销还原全部转曲
                    log("几何采集完成，已撤销临时转曲")
                except Exception as exc:
                    log("撤销临时转曲失败（对象会保持曲线形态）：%r" % (exc,))
        return geo

    def _get_by_addr(self, shapes, addr):
        try:
            obj = shapes
            for k in addr[:-1]:
                obj = obj.Item(k).Shapes
            return obj.Item(addr[-1])
        except Exception:
            return None

    def scan_lines(self, mode, thr_mm, ref_layers=None, progress=None, force=False,
                   eps_mm=None):
        if linecheck is None:
            return None, "缺少检测模块 linecheck.py"
        if self._line_geo is None or force:
            self._line_geo = self.build_line_geo(progress)
        if not self._line_geo:
            return None, "没有可检测的线条"
        refs = list(ref_layers or [])
        if not refs:
            refs = [n for n in self._line_geo.keys()]     # 兜底：全部图层当参考
        kw = {}
        if eps_mm:
            kw["eps_mm"] = float(eps_mm)
        issues = linecheck.analyze(self._line_geo, mode, thr_mm, boundary_layers=refs, **kw)
        # 智能填色的原理：曲线必须真正相交/接触才成「面」。
        # 所以真正导致溢出的只有一种情况——线的端头悬空（没接到任何线上），
        # 填色从端头绕过去，把两侧并成一片。这里按「端头」找，一处只报一次。
        gaps = []
        self._faces_n = 0
        self._graph = None
        try:
            if facemod is not None:
                st = facemod.graph_stats(self._line_geo)
                self._graph = st
                gaps = facemod.dangling_ends(
                    self._line_geo, tol_mm=thr_mm,
                    eps_mm=float(eps_mm or linecheck.EPS_MM), stats=st)
                # 逐条模拟：只有真能让「可填面」+1 的才算有效
                if gaps:
                    facemod.validate_repairs(self._line_geo, gaps)
                log("拓扑: F=%d 面, 悬空 %d, 容差内 %d"
                    % (st["F"], st["dangling"], len(gaps)))
            else:
                gaps = linecheck.analyze_open_ends(
                    self._line_geo, tol_mm=thr_mm,
                    eps_mm=float(eps_mm or linecheck.EPS_MM))
        except Exception as exc:
            log("端点检测失败: %r" % (exc,))
            gaps = []
        # 只报「悬空端头」：一类问题、一处一次，不再叠加重叠的点
        allx = list(gaps) if gaps else list(issues)
        allx.sort(key=lambda x: (x.get("kind") != "gap", x["dist_mm"]))
        for i, it in enumerate(allx, 1):
            it["no"] = i
        return allx, ""

    def layer_names(self):
        out = []
        try:
            page = self.doc.ActivePage
            for i in range(1, int(page.Layers.Count) + 1):
                out.append(str(page.Layers.Item(i).Name))
        except Exception:
            pass
        return out

    def _resolve(self, layer_name, addr):
        try:
            L = self.doc.ActivePage.Layers.Item(layer_name)
        except Exception:
            return None
        obj = L.Shapes
        for k in addr[:-1]:
            try:
                obj = obj.Item(k).Shapes
            except Exception:
                return None
        try:
            return obj.Item(addr[-1])
        except Exception:
            return None

    def extend_issue(self, it):
        """按「端点 → 前一个节点」方向延长到交点。

        不依赖子路径序号：直接按端点坐标在对象上找对应的那一端，
        找不到就放弃（说明几何变了），绝不乱改。
        """
        if not it.get("extend_to"):
            return False, "这一处没有可用的延长方案"
        sh = self._resolve(it["layer"], it["addr"])
        if sh is None:
            return False, "找不到原线条（可能已被移动或删除）"
        try:
            if int(sh.Type) != 3:
                sh.ConvertToCurves()
                sh = self._resolve(it["layer"], it["addr"])
                if sh is None:
                    return False, "转曲后找不到对象"
            c = sh.Curve
        except Exception as exc:
            return False, "读取曲线失败：%s" % exc

        tx = self.mm_to_doc(it["pt"][0])
        ty = self.mm_to_doc(it["pt"][1])
        tol = self.mm_to_doc(0.02)
        best = None
        try:
            n = int(c.SubPaths.Count)
        except Exception:
            n = 0
        for i in range(1, n + 1):
            try:
                sp = c.SubPaths.Item(i)
                if bool(sp.Closed):
                    continue
                cnt = int(sp.Nodes.Count)
                for end, idx in ((0, 1), (1, cnt)):
                    nd = sp.Nodes.Item(idx)
                    d = math.hypot(nd.PositionX - tx, nd.PositionY - ty)
                    if d <= tol and (best is None or d < best[0]):
                        best = (d, sp, end)
            except Exception:
                continue
        if best is None:
            return False, "这条线上找不到对应端点（几何已变），已跳过"

        _, sp, end = best
        try:
            x = self.mm_to_doc(it["extend_to"][0])
            y = self.mm_to_doc(it["extend_to"][1])
            if end == 1:
                sp.AppendLineSegment(x, y)
            else:
                sp.AppendLineSegment(x, y, True)
        except Exception as exc:
            return False, "延长失败：%s" % exc

        try:
            for s in self._line_geo.get(it["layer"], []):
                if s["addr"] == it["addr"] and s["sub"] == it.get("sub"):
                    if it["end"] == 1:
                        s["pts"].append(tuple(it["extend_to"]))
                    else:
                        s["pts"].insert(0, tuple(it["extend_to"]))
                    break
        except Exception:
            pass
        if it.get("mode") == "perp":
            return True, "方向偏了，已用垂线补齐 %.2f mm" % (it.get("extend_len_mm") or 0)
        return True, "已延长 %.2f mm" % (it.get("extend_len_mm") or 0)

    # ---- 线段近缝：插节点 + 建桥接线 ---------------------------

    def _polyline_pts(self, sh, sub_idx):
        c = sh.Curve
        sp = c.SubPaths.Item(sub_idx + 1)
        pts = []
        try:
            info = sp.GetCurveInfo()
            for nd in info:
                pts.append((float(nd.PositionX), float(nd.PositionY)))
        except Exception:
            nodes = sp.Nodes
            for k in range(1, int(nodes.Count) + 1):
                nd = nodes.Item(k)
                pts.append((float(nd.PositionX), float(nd.PositionY)))
        return sp, pts

    @staticmethod
    def _offset_to_point(pts, pt):
        """沿折线累计长度，返回最近投影点的 (距离, 绝对偏移)。"""
        best = None
        off = 0.0
        for i in range(len(pts) - 1):
            a, b = pts[i], pts[i + 1]
            dx, dy = b[0] - a[0], b[1] - a[1]
            l2 = dx * dx + dy * dy
            seg_len = math.sqrt(l2) if l2 > 0 else 0.0
            if l2 <= 1e-18:
                continue
            t = ((pt[0] - a[0]) * dx + (pt[1] - a[1]) * dy) / l2
            t = 0.0 if t < 0 else (1.0 if t > 1 else t)
            px, py = a[0] + t * dx, a[1] + t * dy
            d = math.hypot(pt[0] - px, pt[1] - py)
            if best is None or d < best[0]:
                best = (d, off + t * seg_len)
            off += seg_len
        return best if best else (1e9, 0.0)

    def _insert_node(self, sh, sub_idx, pt_mm):
        """在指定点插一个节点（几何不变）。已在该处则不重复插。"""
        if int(sh.Type) != 3:
            sh.ConvertToCurves()
        sp, pts = self._polyline_pts(sh, sub_idx)
        tgt = (self.mm_to_doc(pt_mm[0]), self.mm_to_doc(pt_mm[1]))
        near = self.mm_to_doc(0.002)
        for p in pts:
            if math.hypot(p[0] - tgt[0], p[1] - tgt[1]) <= near:
                return False
        _d, off = self._offset_to_point(pts, tgt)
        sp.AddNodeAt(off, 0)
        return True

    def _move_node_to(self, sh, sub_idx, old_pt, new_pt):
        """把（刚插入的）节点从 old_pt 挪到 new_pt。"""
        c = sh.Curve
        sp = c.SubPaths.Item(sub_idx + 1)
        tgt = (self.mm_to_doc(old_pt[0]), self.mm_to_doc(old_pt[1]))
        nodes = sp.Nodes
        best = None
        for i in range(1, int(nodes.Count) + 1):
            nd = nodes.Item(i)
            d = math.hypot(nd.PositionX - tgt[0], nd.PositionY - tgt[1])
            if best is None or d < best[0]:
                best = (d, nd)
        if best is None or best[0] > self.mm_to_doc(0.05):
            raise RuntimeError("没找到刚插入的节点")
        best[1].SetPosition(self.mm_to_doc(new_pt[0]), self.mm_to_doc(new_pt[1]))

    def _extend_end_to_line(self, sh, sub_idx, other_pt, max_mm=0.5):
        """把 sh 上离 other_pt 最近的那个端点，直接延长到 other_pt（同一条轨迹上追加）。"""
        try:
            if int(sh.Type) != 3:
                sh.ConvertToCurves()
            sp = sh.Curve.SubPaths.Item(sub_idx + 1)
            if bool(sp.Closed):
                return None
            n = int(sp.Nodes.Count)
            best = None
            for end, idx in ((1, n), (0, 1)):
                nd = sp.Nodes.Item(idx)
                d = math.hypot(nd.PositionX - other_pt[0], nd.PositionY - other_pt[1])
                if best is None or d < best[0]:
                    best = (d, end)
            if best is None or best[0] > self.mm_to_doc(max_mm):
                return None
            if best[0] <= self.mm_to_doc(0.001):
                return 0.0
            if best[1] == 1:
                sp.AppendLineSegment(other_pt[0], other_pt[1])
            else:
                sp.AppendLineSegment(other_pt[0], other_pt[1], True)
            return self.doc_to_mm(best[0])
        except Exception as exc:
            log("延长线端失败: %r" % (exc,))
            return None

    def fix_gap(self, it):
        """近缝修复：优先延长线端，其次节点吸附。"""
        pa = (self.mm_to_doc(it["a_pt"][0]), self.mm_to_doc(it["a_pt"][1]))
        pb = (self.mm_to_doc(it["b_pt"][0]), self.mm_to_doc(it["b_pt"][1]))
        a_sh = self._resolve(it["a_layer"], it["a_addr"])
        if a_sh is not None:
            r = self._extend_end_to_line(a_sh, it["a_sub"], pb)
            if r is not None:
                return True, "已延长线端 %.3f mm" % (it.get("dist_mm") or 0)
        b_sh = self._resolve(it["b_layer"], it["b_addr"])
        if b_sh is not None:
            r = self._extend_end_to_line(b_sh, it["b_sub"], pa)
            if r is not None:
                return True, "已延长线端 %.3f mm" % (it.get("dist_mm") or 0)
        return self.snap_issue(it)

    def snap_issue(self, it):
        """近缝修复：不新增对象、不断开轨迹。

        做法：在被吸附线上插一个节点 → 把这个节点挪到对面线上的最近点，
        同时在对面线上的同一坐标也插一个节点。结果两条线共享一个顶点。
        """
        sh = self._resolve(it["snap_layer"], it["snap_addr"])
        if sh is None:
            return False, "找不到被吸附线条"
        if it["snap_on"] == "a":
            other_layer, other_addr, other_sub = it["b_layer"], it["b_addr"], it["b_sub"]
        else:
            other_layer, other_addr, other_sub = it["a_layer"], it["a_addr"], it["a_sub"]
        sh2 = self._resolve(other_layer, other_addr)
        if sh2 is None:
            return False, "找不到对面线条"

        try:
            self._insert_node(sh, it["snap_sub"], it["snap_from"])
        except Exception as exc:
            return False, "插节点失败：%s" % exc
        try:
            self._move_node_to(sh, it["snap_sub"], it["snap_from"], it["snap_to"])
        except Exception as exc:
            return False, "移动节点失败：%s" % exc
        try:
            self._insert_node(sh2, other_sub, it["snap_to"])
        except Exception as exc:
            log("对面线插节点失败（不影响闭合）: %r" % (exc,))
        return True, "已吸附 %.3f mm" % (it.get("dist_mm") or 0)

    def select_shape(self, it):
        if it.get("kind") == "gap":
            ok = False
            for lay, addr in ((it["a_layer"], it["a_addr"]), (it["b_layer"], it["b_addr"])):
                sh = self._resolve(lay, addr)
                if sh is None:
                    continue
                try:
                    if not ok:
                        sh.CreateSelection()
                        ok = True
                    else:
                        sh.AddToSelection()
                except Exception:
                    pass
            return ok
        sh = self._resolve(it["layer"], it["addr"])
        if sh is None:
            return False
        try:
            sh.CreateSelection()
            return True
        except Exception:
            return False

    def clear_marks(self):
        try:
            L = self.doc.ActivePage.Layers.Item(self.MARK_LAYER)
        except Exception:
            return
        try:
            L.Delete()
        except Exception:
            pass

    def mark_issues(self, issues, radius_mm=1.2, progress=None):
        """在「检测标记」图层给每个端点画一个小红圈。"""
        self.clear_marks()
        try:
            layer = self.doc.ActivePage.CreateLayer(self.MARK_LAYER)
        except Exception as exc:
            log("建标记图层失败: %r" % (exc,))
            return 0
        r = self.mm_to_doc(radius_mm)
        w = self.mm_to_doc(0.25)
        n = 0
        total = max(1, len(issues))
        for idx, it in enumerate(issues):
            if progress:
                progress(idx / float(total), "标注端点 %d/%d" % (idx + 1, len(issues)))
            try:
                cx = self.mm_to_doc(it["pt"][0])
                cy = self.mm_to_doc(it["pt"][1])
                el = layer.CreateEllipse2(cx, cy, r, r)
                try:
                    el.Outline.Width = w
                except Exception:
                    pass
                try:
                    # 荧光黄填充 + 65% 透明：看得见但不挡线；描边用红圈
                    c = el.Outline.Color
                    c.HexValue = "#EFFF00"
                    el.Fill.ApplyUniformFill(c)
                    el.Transparency.ApplyUniformTransparency(65)
                    el.Outline.Color.HexValue = "#E60000"
                except Exception:
                    pass
                n += 1
            except Exception:
                pass
        return n

    # ---- 批处理备份 / 同色归并 --------------------------------

    def backup_before_batch(self, tag=""):
        try:
            p = str(self.doc.FullFileName)
        except Exception:
            p = ""
        if not p or not os.path.isfile(p):
            return None
        dest = backup_file(p)
        if dest:
            log("批处理备份(%s) -> %s" % (tag, dest))
        return dest

    def restore_last_backup(self):
        try:
            p = str(self.doc.FullFileName)
        except Exception:
            p = ""
        if not p:
            return False, "当前文档还没保存过，没有备份可回滚"
        files = list_backups(p)
        if not files:
            return False, "这个文件还没有批处理备份"
        src = files[-1]
        try:
            self.doc.Close()           # COM 下直接丢弃改动，不弹保存框（实测）
        except Exception as exc:
            return False, "关闭当前文档失败：%s" % exc
        try:
            shutil.copy2(src, p)
            doc = self.app.OpenDocument(p)
            doc.Activate()
            self.attach()
            return True, "已回滚到备份：%s" % os.path.basename(src)
        except Exception as exc:
            return False, "回滚失败：%s" % exc

    def merge_by_fill(self, progress=None):
        """按填充色归并：每种颜色一个图层，名字取颜色名（重名加序号）。"""
        if colorname is None:
            return None, "缺少颜色映射模块 colors.py"
        page = self.doc.ActivePage
        groups = {}
        skipped = 0
        for li in range(1, int(page.Layers.Count) + 1):
            L = page.Layers.Item(li)
            if str(L.Name) == self.MARK_LAYER:
                continue
            for addr, sh in self._walk(L.Shapes):
                try:
                    if int(sh.Type) in (5, 6):        # 文本 / 位图
                        skipped += 1
                        continue
                    f = sh.Fill
                    if int(f.Type) == 0:              # 无填充（透明）
                        skipped += 1
                        continue
                    hx = str(f.UniformColor.HexValue).upper()
                except Exception:
                    skipped += 1
                    continue
                groups.setdefault(hx, []).append(sh)
        if not groups:
            return None, "没有找到带填充色的图形"

        used = {}
        made = []
        order = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
        for gi, (hx, items) in enumerate(order):
            if progress:
                progress(gi / float(max(1, len(order))),
                         "归并颜色 %d/%d  %s" % (gi + 1, len(order), hx))
            base, _exact = colorname.color_name(hx)
            used[base] = used.get(base, 0) + 1
            lname = base if used[base] == 1 else "%s(%d)" % (base, used[base])
            try:
                layer = page.CreateLayer(lname)
            except Exception as exc:
                log("建图层失败 %s: %r" % (lname, exc))
                continue
            moved = 0
            for sh in items:
                try:
                    sh.MoveToLayer(layer)
                    moved += 1
                except Exception:
                    pass
            made.append((lname, hx, moved))
        return made, ""

    def _write_fill(self, st, rgb):
        """填充色只能通过 JSON 字符串写入，并跟随现有颜色模型。"""
        model = "RGB255"
        try:
            cur = self._fill_json(st)
            sample = str(cur.get("primaryColor", ""))
            if sample.upper().startswith("CMYK"):
                model = "CMYK"
        except Exception:
            pass

        if model == "CMYK":
            c, m, y, k = rgb_to_cmyk(rgb)
            token = "CMYK,USER,%d,%d,%d,%d,100,%s" % (c, m, y, k, NULL_GUID)
        else:
            token = "RGB255,USER,%d,%d,%d,100,%s" % (rgb[0], rgb[1], rgb[2], NULL_GUID)

        payload = json.dumps({
            "primaryColor": token,
            "secondaryColor": "RGB255,USER,255,255,255,100," + NULL_GUID,
            "type": "1",
        })
        st.SetPropertyAsString("fill", payload)


# ---------------------------------------------------------------- 进度窗

class Cancelled(Exception):
    """长任务被用户取消。"""


def ask_centered(root, title, lines, ok_text="确定", cancel_text="稍后再说"):
    """屏幕正中央的确认框（标题栏 + 说明 + 两个按钮），返回 True/False。"""
    res = {"v": False}
    top = tk.Toplevel(root)
    top.title(title)
    top.configure(bg=BG)
    top.attributes("-topmost", True)
    top.resizable(False, False)

    body = tk.Frame(top, bg=BG)
    body.pack(fill="both", expand=True, padx=18, pady=(16, 6))
    first = True
    for line in lines:
        if not line:
            continue
        tk.Label(body, text=line, bg=BG, fg=TXT if first else MUTED,
                 font=FONT_B if first else FONT_S, anchor="w", justify="left",
                 wraplength=380).pack(fill="x", pady=(0, 10) if first else (0, 4))
        first = False

    row = tk.Frame(top, bg=BG)
    row.pack(fill="x", padx=18, pady=(4, 14))

    def close(v):
        res["v"] = v
        try:
            top.grab_release()
        except Exception:
            pass
        top.destroy()

    b_ok = tk.Label(row, text=ok_text, bg=ACCENT, fg="#FFFFFF", font=FONT_S,
                    padx=16, pady=6, cursor="hand2")
    b_ok.pack(side="right")
    b_ok.bind("<Button-1>", lambda e: close(True))
    b_no = tk.Label(row, text=cancel_text, bg=CHIP_BG, fg=TXT, font=FONT_S,
                    padx=16, pady=6, cursor="hand2")
    b_no.pack(side="right", padx=(0, 8))
    b_no.bind("<Button-1>", lambda e: close(False))
    top.protocol("WM_DELETE_WINDOW", lambda: close(False))

    top.update_idletasks()
    w = max(420, top.winfo_reqwidth())
    h = top.winfo_reqheight()
    sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
    top.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 2)))
    try:
        top.grab_set()
        top.focus_force()
    except Exception:
        pass
    root.wait_window(top)
    return res["v"]


class ProgressDialog(object):
    """一条进度条 + 百分比 + 取消按钮；取消只是置位，由任务在安全点检查。"""

    def __init__(self, app, title, head=""):
        self.app = app
        self._cancel = False
        self._closed = False
        top = tk.Toplevel(app.root)
        self.top = top
        top.title(title)
        top.configure(bg=BG)
        top.attributes("-topmost", True)
        top.resizable(False, False)
        W_, H_ = 380, 146
        sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
        top.geometry("%dx%d+%d+%d" % (W_, H_, max(0, (sw - W_) // 2),
                                      max(0, (sh - H_) // 2)))
        top.protocol("WM_DELETE_WINDOW", self.cancel)

        tk.Label(top, text=title, bg=BG, fg=TXT, font=FONT_B).pack(
            anchor="w", padx=16, pady=(14, 2))
        self.lb = tk.Label(top, text=(head or "准备中…"), bg=BG, fg=MUTED,
                           font=FONT_S, anchor="w")
        self.lb.pack(fill="x", padx=16)

        self.cv = tk.Canvas(top, height=10, bg=LINE, highlightthickness=0)
        self.cv.pack(fill="x", padx=16, pady=(12, 6))
        self.bar = self.cv.create_rectangle(0, 0, 0, 10, fill=ACCENT, outline="")
        self.pct = tk.Label(top, text="0%", bg=BG, fg=MUTED, font=FONT_S)
        self.pct.pack(anchor="e", padx=16)

        self.btn = tk.Label(top, text="取消", bg=CHIP_BG, fg=TXT, font=FONT_S,
                            padx=16, pady=4, cursor="hand2")
        self.btn.pack(pady=(4, 10))
        self.btn.bind("<Button-1>", lambda e: self.cancel())
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=CHIP_HOVER))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=CHIP_BG))
        try:
            top.update()
        except Exception:
            pass

    def set(self, frac, text=None):
        if self._closed:
            return
        try:
            frac = 0.0 if frac < 0 else (1.0 if frac > 1 else frac)
            w = self.cv.winfo_width()
            if w <= 1:
                w = 328
            self.cv.coords(self.bar, 0, 0, int(w * frac), 10)
            self.pct.config(text="%d%%" % int(frac * 100))
            if text is not None:
                self.lb.config(text=text)
            self.top.update()
        except Exception:
            pass

    def cancel(self):
        if not self._cancel:
            self._cancel = True
            try:
                self.btn.config(text="正在取消…", bg=TAG_BG)
                self.lb.config(text="正在收尾…")
                self.top.update()
            except Exception:
                pass

    def cancelled(self):
        return self._cancel

    def check(self):
        if self._cancel:
            raise Cancelled()

    def close(self):
        self._closed = True
        try:
            self.top.destroy()
        except Exception:
            pass


# ---------------------------------------------------------------- 界面

class FloatWindow(object):

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CorelDRAW 助手")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=LINE)

        self.cfg = self._load_cfg()
        self.bridge = CdrBridge()
        self.collapsed = False
        self._drag = None
        self._pending = None
        self._auto_collapsed = False
        self._ever_connected = False
        self._switching = False

        self.doc_list = []
        self.current_key = None
        self.doc_states = {}
        self.pal_cache = {}
        self._doc_items = []

        self._update_checked = False   # 每次启动只自动查一次
        self._update_manual = False
        self._update_done = False      # 后台线程写、_tick 在主线程读
        self._update_info = None
        self._update_latest = None     # 线上比本地新的那份 info（有值 = 有小红点）
        self._status_hold = 0          # 状态栏“别被顶掉”的截止时间
        self._line_geo = None          # 线条几何缓存（按文档）
        self._backup_track = {}        # 路径 -> 备份时刻（用于"保存并关闭"后清缓存）
        try:
            self._act_seen = os.path.getmtime(ACTIVATE_FLAG) if os.path.exists(ACTIVATE_FLAG) else 0
        except Exception:
            self._act_seen = 0
        self._busy = False             # 长操作期间暂停心跳，避免 COM 重入

        self.state = {
            "width_mm": self.cfg.get("width_mm", 0.35),
            "outline_rgb": tuple(self.cfg.get("outline_rgb", (0, 0, 0))),
            "fill_mode": self.cfg.get("fill_mode", "none"),
            "fill_rgb": tuple(self.cfg.get("fill_rgb", (217, 217, 218))),
            "kinds": {"graphic": True, "artistic": False},
            "live": self.cfg.get("live", True),
            "line_caps": self.cfg.get("line_caps", 0),
            "line_join": self.cfg.get("line_join", 0),
            "justify": self.cfg.get("justify", 0),
            "scale_with": self.cfg.get("scale_with", False),
            "behind_fill": self.cfg.get("behind_fill", False),
        }
        self.adv_open = bool(self.cfg.get("adv_open", False))
        self.linestyle = int(self.cfg.get("linestyle", 2))

        self._build()
        self._place()
        self._init_tray()
        self.root.after(250, self._connect)
        self.root.after(900, self._tick)
        self.root.after(3000, self._check_update)
        self._old_clean_tries = 0

    # ---- 配置 -------------------------------------------------

    def _load_cfg(self):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_cfg(self):
        cfg = {
            "x": self.root.winfo_x(), "y": self.root.winfo_y(),
            "width_mm": self.state["width_mm"],
            "outline_rgb": list(self.state["outline_rgb"]),
            "fill_mode": self.state["fill_mode"],
            "fill_rgb": list(self.state["fill_rgb"]),
            "live": self.state["live"],
            "line_caps": self.state.get("line_caps", 0),
            "line_join": self.state.get("line_join", 0),
            "justify": self.state.get("justify", 0),
            "scale_with": self.state.get("scale_with", False),
            "behind_fill": self.state.get("behind_fill", False),
            "adv_open": self.adv_open,
            "linestyle": self.linestyle,
        }
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _place(self):
        sw = self.root.winfo_screenwidth()
        x = int(self.cfg.get("x", sw - W - 48))
        y = int(self.cfg.get("y", 100))
        if x < -W or x > sw - 40:
            x = sw - W - 48
        if y < 0 or y > self.root.winfo_screenheight() - 60:
            y = 100
        h = self._full_height()
        self.root.geometry("%dx%d+%d+%d" % (W, h, x, y))

    def _full_height(self):
        try:
            self.root.update_idletasks()
            need = self.body.winfo_reqheight() + BAR_H + QUICK_H + VER_H + 2
            return max(H, need)
        except Exception:
            return H + 130

    # ---- 组件 -------------------------------------------------

    def _label(self, parent, text, fg=TXT, font=FONT, **kw):
        return tk.Label(parent, text=text, bg=BG, fg=fg, font=font, **kw)

    def _chip(self, parent, text, cmd, width=None):
        lb = tk.Label(parent, text=text, bg=CHIP_BG, fg=TXT, font=FONT_S,
                      padx=6, pady=5, cursor="hand2", width=width or 0)
        lb.bind("<Button-1>", lambda e: cmd())
        lb.bind("<Enter>", lambda e: lb.config(bg=CHIP_HOVER))
        lb.bind("<Leave>", lambda e: lb.config(bg=CHIP_BG))
        return lb

    # ---- 构建 -------------------------------------------------

    def _build(self):
        self.shell = tk.Frame(self.root, bg=BG)
        self.shell.pack(fill="both", expand=True, padx=1, pady=1)

        self._build_bar()

        # 固定底栏：先占住底部，永远贴底
        self._build_quickbar(self.shell)
        self._build_verbar(self.shell)

        self.body = tk.Frame(self.shell, bg=BG)
        self.body.pack(fill="both", expand=True)

        self._build_docbar(self.body)

        pad = tk.Frame(self.body, bg=BG)
        pad.pack(fill="both", expand=True, padx=14, pady=(6, 12))

        # ① 默认样式（决定"接下来画的线"长什么样）
        self._build_width(pad)
        self._sep(pad)
        self._build_color(pad)
        self._sep(pad)
        self._build_fill(pad)
        self._sep(pad)
        self._build_targets(pad)
        self._build_advanced(pad)
        self._sep(pad)

        # ② 应用到文档（紧跟在设置下面）
        self._build_apply(pad)
        self._sep(pad)

        # ③ 批量改整份文档
        self._build_batch(pad)
        self._sep(pad)

        # ④ 检查与修补
        self._build_linescan(pad)

    def _build_apply(self, parent):
        self.btn_apply = tk.Label(parent, text="应用到 CorelDRAW", bg=ACCENT, fg="#FFFFFF",
                                  font=FONT_B, pady=7, cursor="hand2")
        self.btn_apply.pack(fill="x")
        self.btn_apply.bind("<Button-1>", lambda e: self._apply(explicit=True))
        self.btn_apply.bind("<Enter>", lambda e: self.btn_apply.config(bg=ACCENT_HOVER))
        self.btn_apply.bind("<Leave>", lambda e: self.btn_apply.config(bg=ACCENT))

        self.status = tk.Label(parent, text="", bg=BG, fg=MUTED, font=FONT_S, anchor="w")
        self.status.pack(fill="x", pady=(6, 0))

    def _init_tray(self):
        """系统托盘图标：单击/双击显示并居中，右键有菜单。"""
        self.tray = None
        if not HAS_TRAY:
            log("托盘不可用: %s" % globals().get("_TRAY_ERR", "tray 模块缺失"))
            return
        try:
            self.tray = traymod.TrayIcon(
                asset_path("app.ico"), "CorelDRAW 助手",
                on_show=lambda: self.root.after(0, self._show_from_tray),
                on_check=lambda: self.root.after(
                    0, lambda: self._check_update(manual=True)),
                on_quit=lambda: self.root.after(0, self._quit))
            self.tray.start()
        except Exception as exc:
            log("托盘启动失败: %r" % (exc,))
            self.tray = None

    def _show_from_tray(self):
        try:
            if self.collapsed:
                self._toggle_collapse()
            self.root.deiconify()
            self.root.attributes("-topmost", True)
            self._center_self()
            self.root.lift()
            self.root.focus_force()
            self.set_status("已显示窗口", "ok", hold=3)
        except Exception as exc:
            log("从托盘显示失败: %r" % (exc,))

    def _open_feedback(self):
        """反馈 / 许愿：发到 sjck.kust.work 的留言板。"""
        top = tk.Toplevel(self.root)
        top.title("反馈 / 许愿")
        top.configure(bg=BG)
        top.attributes("-topmost", True)
        top.resizable(False, False)

        pad = tk.Frame(top, bg=BG)
        pad.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(pad, text="问题反馈 / 功能许愿", bg=BG, fg=TXT, font=FONT_B).pack(anchor="w")
        tk.Label(pad, text="会发到 sjck.kust.work 的留言板，我看到就会处理。",
                 bg=BG, fg=MUTED, font=FONT_S).pack(anchor="w", pady=(2, 10))

        kinds = {"v": "bug"}
        krow = tk.Frame(pad, bg=BG)
        krow.pack(fill="x", pady=(0, 8))
        kb = {}

        def paint_kind():
            for k, w in kb.items():
                on = kinds["v"] == k
                w.config(bg=ACCENT_SOFT if on else CHIP_BG,
                         fg=ACCENT if on else MUTED)

        for k, label in (("bug", "问题反馈"), ("wish", "功能许愿")):
            w = tk.Label(krow, text=label, font=FONT_S, padx=12, pady=5, cursor="hand2")
            w.pack(side="left", padx=(0, 8))
            kb[k] = w

            def pick(e, key=k):
                kinds["v"] = key
                paint_kind()
            w.bind("<Button-1>", pick)
        paint_kind()

        nrow = tk.Frame(pad, bg=BG)
        nrow.pack(fill="x", pady=(0, 8))
        tk.Label(nrow, text="称呼", bg=BG, fg=MUTED, font=FONT_S).pack(side="left")
        name_var = tk.StringVar(value="")
        tk.Entry(nrow, textvariable=name_var, font=FONT_S, relief="flat", bg=FIELD_BG,
                 fg=TXT, insertbackground=TXT, highlightthickness=1,
                 highlightbackground=LINE, highlightcolor=ACCENT).pack(
            side="left", fill="x", expand=True, padx=(8, 0), ipady=3)

        txt = tk.Text(pad, height=7, font=FONT_S, relief="flat", bg=FIELD_BG, fg=TXT,
                      insertbackground=TXT, highlightthickness=1,
                      highlightbackground=LINE, highlightcolor=ACCENT, wrap="word")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", "")
        txt.focus_set()

        st = tk.Label(pad, text="", bg=BG, fg=MUTED, font=FONT_S, anchor="w")
        st.pack(fill="x", pady=(6, 0))

        row = tk.Frame(pad, bg=BG)
        row.pack(fill="x", pady=(8, 0))

        def close():
            top.destroy()

        def send():
            body = txt.get("1.0", "end").strip()
            if not body:
                st.config(text="先写点内容吧", fg=BAD)
                return
            btn_send.config(state="disabled", text="发送中…")
            st.config(text="正在发送…", fg=MUTED)

            def worker():
                err = ""
                try:
                    payload = json.dumps({
                        "text": body, "name": name_var.get().strip(),
                        "kind": kinds["v"], "version": APP_VERSION,
                    }).encode("utf-8")
                    req = urllib.request.Request(
                        FEEDBACK_URL, data=payload,
                        headers={"Content-Type": "application/json",
                                 "User-Agent": "CDRHelper/%s" % APP_VERSION})
                    with urllib.request.urlopen(req, timeout=15) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                    if not data.get("ok"):
                        err = data.get("error") or "服务器拒绝了这条留言"
                except Exception as exc:
                    err = "%s: %s" % (type(exc).__name__, str(exc)[:60])

                def done():
                    try:
                        if err:
                            btn_send.config(state="normal", text="发送")
                            st.config(text="发送失败 · %s" % err, fg=BAD)
                        else:
                            st.config(text="已发送，谢谢！", fg=OK)
                            txt.delete("1.0", "end")
                            btn_send.config(state="normal", text="再发一条")
                            self.set_status("反馈已发送", "ok", hold=5)
                    except Exception:
                        pass

                try:
                    self.root.after(0, done)
                except Exception:
                    pass

            threading.Thread(target=worker, daemon=True).start()

        def open_board():
            try:
                webbrowser.open(FEEDBACK_URL)
            except Exception:
                pass

        btn_send = tk.Label(row, text="发送", bg=ACCENT, fg="#FFFFFF", font=FONT_S,
                            padx=18, pady=6, cursor="hand2")
        btn_send.pack(side="right")
        btn_send.bind("<Button-1>", lambda e: send())
        tk.Label(row, text="看看留言板", bg=CHIP_BG, fg=TXT, font=FONT_S, padx=12, pady=6,
                 cursor="hand2").pack(side="right", padx=(0, 8))
        row.winfo_children()[-1].bind("<Button-1>", lambda e: open_board())
        tk.Label(row, text="取消", bg=CHIP_BG, fg=TXT, font=FONT_S, padx=12, pady=6,
                 cursor="hand2").pack(side="left")
        row.winfo_children()[-1].bind("<Button-1>", lambda e: close())

        top.protocol("WM_DELETE_WINDOW", close)
        top.update_idletasks()
        w, h = max(430, top.winfo_reqwidth()), max(360, top.winfo_reqheight())
        sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
        top.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, (sh - h) // 2))

    def _build_verbar(self, parent):
        bar = tk.Frame(parent, bg=BG, height=VER_H)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self.verbar = bar
        tk.Frame(bar, bg=LINE, height=1).pack(side="top", fill="x")

        row = tk.Frame(bar, bg=BG)
        row.pack(fill="both", expand=True)

        self.ver_dot = tk.Canvas(row, width=10, height=10, bg=BG, highlightthickness=0)
        self.ver_dot.pack(side="left", padx=(14, 3))
        self.ver_dot_id = self.ver_dot.create_oval(2, 2, 8, 8, fill=BAD, outline="")
        self.ver_dot.itemconfig(self.ver_dot_id, state="hidden")

        self.ver_lb = tk.Label(row, text="版本 v%s" % APP_VERSION, bg=BG, fg=MUTED,
                               font=FONT_S, cursor="hand2")
        self.ver_lb.pack(side="left")

        tk.Label(row, text="内测版 · 开发者：黄双成", bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 7)).pack(side="left", padx=(6, 0))
        self.ver_lb.bind("<Button-1>", lambda e: self._on_ver_click())

        self.ver_btn = tk.Label(row, text="⟳", bg=BG, fg=MUTED, font=FONT, padx=8,
                                cursor="hand2")
        self.ver_btn.pack(side="right", padx=(0, 6))
        self.ver_btn.bind("<Button-1>", lambda e: self._check_update(manual=True))
        self.ver_btn.bind("<Enter>", lambda e: self.ver_btn.config(fg=ACCENT))
        self.ver_btn.bind("<Leave>", lambda e: self.ver_btn.config(fg=MUTED))

    def _set_ver_state(self, text, level="info", has_new=False):
        try:
            self.ver_lb.config(text=text,
                               fg={"info": MUTED, "ok": OK, "warn": WARN,
                                   "bad": BAD, "new": ACCENT}.get(level, MUTED))
            self.ver_dot.itemconfig(self.ver_dot_id,
                                    state="normal" if has_new else "hidden")
        except Exception:
            pass

    def _on_ver_click(self):
        if self._update_latest:
            self._prompt_update(self._update_latest)
        else:
            self._check_update(manual=True)

    def _build_quickbar(self, parent):
        bar = tk.Frame(parent, bg=TAG_BG, height=QUICK_H)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self.quickbar = bar

        def mk(text, cmd):
            lb = tk.Label(bar, text=text, bg=TAG_BG, fg=TXT, font=FONT_S, cursor="hand2")
            lb.pack(side="left", fill="both", expand=True)
            lb.bind("<Button-1>", lambda e: cmd())
            lb.bind("<Enter>", lambda e, x=lb: x.config(bg=CHIP_HOVER))
            lb.bind("<Leave>", lambda e, x=lb: x.config(bg=TAG_BG))
            return lb

        self.btn_undo = mk("⟲  撤销", self._cdr_undo)
        tk.Frame(bar, bg=LINE, width=1).pack(side="left", fill="y")
        self.btn_redo = mk("⟳  重做", self._cdr_redo)
        tk.Frame(bar, bg=LINE, width=1).pack(side="left", fill="y")
        self.btn_clear = mk("⌫  清空缓存", self._clear_cache)
        tk.Frame(bar, bg=LINE, width=1).pack(side="left", fill="y")
        self.btn_fb = mk("✎  反馈", self._open_feedback)

    def _build_advanced(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", pady=(8, 0))
        tk.Frame(bar, bg=LINE, height=1).pack(fill="x")

        self.adv_lb = tk.Label(bar, text="", font=FONT_S, bg=BG, fg=MUTED, cursor="hand2")
        self.adv_lb.pack(pady=(6, 0))
        self.adv_lb.bind("<Button-1>", lambda e: self._toggle_advanced())
        self.adv_bar = bar

        self.adv_frame = tk.Frame(parent, bg=BG)

        r1 = tk.Frame(self.adv_frame, bg=BG)
        r1.pack(fill="x", pady=(7, 0))
        self._label(r1, "线帽", fg=MUTED, font=FONT_S).pack(side="left", padx=(0, 5))
        self.cap_btns = {}
        for v, t in ((0, "平"), (1, "圆"), (2, "方")):
            lb = tk.Label(r1, text=t, font=FONT_S, padx=7, pady=3, cursor="hand2")
            lb.pack(side="left", padx=(0, 3))
            lb.bind("<Button-1>", lambda e, x=v: self._set_adv("line_caps", x))
            self.cap_btns[v] = lb

        self._label(r1, "线角", fg=MUTED, font=FONT_S).pack(side="left", padx=(9, 5))
        self.join_btns = {}
        for v, t in ((0, "尖"), (1, "圆"), (2, "斜")):
            lb = tk.Label(r1, text=t, font=FONT_S, padx=7, pady=3, cursor="hand2")
            lb.pack(side="left", padx=(0, 3))
            lb.bind("<Button-1>", lambda e, x=v: self._set_adv("line_join", x))
            self.join_btns[v] = lb

        r2 = tk.Frame(self.adv_frame, bg=BG)
        r2.pack(fill="x", pady=(6, 0))
        self._label(r2, "轮廓位置", fg=MUTED, font=FONT_S).pack(side="left", padx=(0, 5))
        self.just_btns = {}
        for v, t in ((0, "中"), (1, "内"), (2, "外")):
            lb = tk.Label(r2, text=t, font=FONT_S, padx=7, pady=3, cursor="hand2")
            lb.pack(side="left", padx=(0, 3))
            lb.bind("<Button-1>", lambda e, x=v: self._set_adv("justify", x))
            self.just_btns[v] = lb

        self.scale_lb = tk.Label(r2, text="", font=FONT_S, bg=BG, cursor="hand2")
        self.scale_lb.pack(side="right")
        self.scale_lb.bind("<Button-1>", lambda e: self._toggle_adv_bool("scale_with"))

        r3 = tk.Frame(self.adv_frame, bg=BG)
        r3.pack(fill="x", pady=(6, 0))
        self._label(r3, "线型", fg=MUTED, font=FONT_S).pack(side="left", padx=(0, 5))
        self.ls_btns = {}
        for v, t in ((1, "实"), (2, "虚"), (3, "点"), (4, "点划"), (5, "双点")):
            lb = tk.Label(r3, text=t, font=FONT_S, padx=6, pady=3, cursor="hand2")
            lb.pack(side="left", padx=(0, 3))
            lb.bind("<Button-1>", lambda e, x=v: self._apply_linestyle(x))
            self.ls_btns[v] = lb
        self._label(r3, "改选中对象", fg=MUTED, font=FONT_S).pack(side="right")

        self._paint_advanced()
        if self.adv_open:
            self.adv_frame.pack(fill="x")

    def _paint_advanced(self):
        self.adv_lb.config(text=("更多轮廓参数  ▴" if self.adv_open else "更多轮廓参数  ▾"),
                           fg=ACCENT if self.adv_open else MUTED)
        for key, btns, cur in (
            ("line_caps", self.cap_btns, self.state.get("line_caps")),
            ("line_join", self.join_btns, self.state.get("line_join")),
            ("justify", self.just_btns, self.state.get("justify")),
        ):
            for v, lb in btns.items():
                if v == cur:
                    lb.config(bg=SEL_BG, fg=SEL_FG)
                else:
                    lb.config(bg=CHIP_BG, fg=TXT)
        for v, lb in self.ls_btns.items():
            if v == self.linestyle:
                lb.config(bg=SEL_BG, fg=SEL_FG)
            else:
                lb.config(bg=CHIP_BG, fg=TXT)
        if self.state.get("scale_with"):
            self.scale_lb.config(text="✓ 随对象缩放", fg=ACCENT)
        else:
            self.scale_lb.config(text="○ 随对象缩放", fg=MUTED)

    def _toggle_advanced(self):
        self.adv_open = not self.adv_open
        if self.adv_open:
            self.adv_frame.pack(fill="x")
        else:
            self.adv_frame.pack_forget()
        self.root.update_idletasks()
        if not self.collapsed:
            self.root.geometry("%dx%d" % (W, self._full_height()))
        self._paint_advanced()
        self._save_cfg()

    def _set_adv(self, key, value):
        self.state[key] = value
        self._paint_advanced()
        self._live_apply()

    def _toggle_adv_bool(self, key):
        self.state[key] = not self.state.get(key)
        self._paint_advanced()
        self._live_apply()

    def _apply_linestyle(self, index):
        self.linestyle = index
        self._paint_advanced()
        self._save_cfg()
        if not self.bridge.connected:
            self.set_status("未连接 CorelDRAW", "bad")
            return
        done, msg = self.bridge.apply_linestyle(index)
        self.set_status(msg, "ok" if done else "warn")

    def _group_by_linetype(self):
        """一键把当前文档所有图层的线条按 实线 / 虚线 分成两个组合。"""
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
            self._ever_connected = True
            self.current_key = None
            self._refresh_docs()

        if self.current_key and self.bridge.doc_name != self.current_key:
            for idx, name in self.doc_list:
                if name == self.current_key:
                    self.bridge.use_doc(idx, activate=False)
                    break

        target = self.bridge.doc_name or "当前文档"
        if not messagebox.askyesno(
                "按线型分组",
                "把「%s」所有图层里的线条，按线型分成两组：\n\n"
                "    实线 →「实线」组合\n"
                "    虚线 →「虚线」组合\n\n"
                "· 空图层、锁定的图层会跳过；\n"
                "· 某一类少于两个对象时不建组；\n"
                "· 分组后可在 CorelDRAW 里 Ctrl+Z 一次撤销。\n\n继续？" % target):
            return

        self._backup_for_batch("按线型分组")
        # layer.Activate() 也会泵消息，先把挂起的延迟写入掐掉
        self._cancel_pending()
        self.set_status("正在按线型分组…")
        self.root.update_idletasks()

        dlg = ProgressDialog(self, "按线型分组", "准备中…")
        cancelled = False

        def gprog(f, t):
            dlg.set(f, t)
            dlg.check()

        self._busy = True
        try:
            rows, msg = self.bridge.group_by_linetype(progress=gprog)
        except Cancelled:
            cancelled = True
            rows, msg = None, ""
        finally:
            self._busy = False
            dlg.close()

        if cancelled:
            self.set_status("已取消分组（已处理的图层仍生效，可 Ctrl+Z 撤销）", "warn", hold=8)
            return
        if rows is None:
            self.set_status(msg, "warn")
            messagebox.showwarning("按线型分组", msg)
            return

        self.set_status(msg, "ok")
        lines = []
        for pno, name, s, d, note in rows:
            if s < 2 and d < 2 and not note:
                continue
            txt = "第%d页 · %s：实线 %d / 虚线 %d" % (pno, name, max(s, 0), max(d, 0))
            if note:
                txt += "（%s）" % note
            lines.append(txt)
        detail = "\n".join(lines[:15])
        if len(lines) > 15:
            detail += "\n…… 共 %d 个图层" % len(lines)
        messagebox.showinfo("按线型分组 · 完成", "%s\n\n%s" % (msg, detail))

    # ---- 线条检测 UI -------------------------------------------

    # ---- 批量：同色归并 + 备份/撤销 ----------------------------

    def _cdr_undo(self):
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
        try:
            self.bridge.doc.Undo()
            self._line_geo = None
            self.set_status("已撤销一步（检测缓存已失效，下次检测会重扫）", "ok", hold=5)
        except Exception as exc:
            self.set_status("撤销失败：%s" % str(exc)[:40], "bad", hold=5)

    def _cdr_redo(self):
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
        try:
            self.bridge.doc.Redo()
            self._line_geo = None
            self.set_status("已重做一步（检测缓存已失效，下次检测会重扫）", "ok", hold=5)
        except Exception as exc:
            self.set_status("重做失败：%s" % str(exc)[:40], "bad", hold=5)

    def _clear_cache(self):
        try:
            p = str(self.bridge.doc.FullFileName)
        except Exception:
            p = ""
        n = len(list_backups(p)) if p else 0
        if not messagebox.askyesno(
                "清空缓存",
                "会清掉这些缓存（不影响图形本身）：\n\n"
                "· 这个文件的批处理备份 %d 份\n"
                "· 线条检测的几何缓存\n"
                "· 文档里的「检测标记」图层\n\n继续？" % n):
            return
        if p:
            clear_backups(p)
            self._backup_track.pop(p, None)
        self._line_geo = None
        try:
            self.bridge.clear_marks()
        except Exception:
            pass
        self.set_status("缓存已清空（备份 %d 份已删）" % n, "ok", hold=6)

    def _backup_for_batch(self, tag):
        dest = self.bridge.backup_before_batch(tag)
        if not dest:
            return None
        try:
            p = str(self.bridge.doc.FullFileName)
            self._backup_track[p] = time.time()
        except Exception:
            pass
        try:
            n = len(list_backups(str(self.bridge.doc.FullFileName)))
        except Exception:
            n = 0
        self.set_status("已备份（%d/10）%s" % (n, os.path.basename(dest)), hold=4)
        return dest

    def _run_color_merge(self):
        if self._busy:
            return
        if colorname is None:
            messagebox.showerror("缺少模块", "colors.py 没有加载成功")
            return
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
            self._ever_connected = True
            self.current_key = None
            self._refresh_docs()
        if not self.bridge.ensure_doc():
            self.set_status("没有打开的文档", "warn")
            return
        if not messagebox.askyesno(
                "同色归并",
                "把当前文档里**填充色相同**的图形归到同一个图层，图层名用颜色名。\n\n"
                "· 无填充 / 透明对象跳过\n"
                "· 文本、位图不参与\n"
                "· 同色名不同色时自动加 (1)(2) 区分\n"
                "· 动手前会自动备份（每个文档留最近 10 份）\n\n继续？"):
            return

        self._backup_for_batch("同色归并")
        self._cancel_pending()
        dlg = ProgressDialog(self, "同色归并", "正在统计颜色…")
        cancelled = False

        def mprog(f, t):
            dlg.set(f, t)
            dlg.check()

        self._busy = True
        try:
            made, err = self.bridge.merge_by_fill(progress=mprog)
        except Cancelled:
            cancelled = True
            made, err = None, ""
        finally:
            self._busy = False
            dlg.close()

        if cancelled:
            self.set_status("已取消归并（已建好的图层保留，可 Ctrl+Z 撤销）", "warn", hold=8)
            return

        if made is None:
            self.set_status(err or "归并失败", "bad", hold=8)
            messagebox.showwarning("同色归并", err or "归并失败")
            return
        total = sum(c for _n, _h, c in made)
        self.set_status("同色归并：%d 种颜色 / %d 个图形" % (len(made), total), "ok", hold=10)
        lines = ["%-8s %s  %d 个" % (n, h, c) for n, h, c in made[:18]]
        if len(made) > 18:
            lines.append("…… 共 %d 个颜色图层" % len(made))
        messagebox.showinfo("同色归并 · 完成",
                            "已归并 %d 种填充色，共 %d 个图形。\n\n%s"
                            % (len(made), total, "\n".join(lines)))

    def _undo_batch(self):
        if self._busy:
            return
        try:
            p = str(self.bridge.doc.FullFileName)
        except Exception:
            p = ""
        if not p:
            self.set_status("当前文档还没保存过", "warn")
            return
        files = list_backups(p)
        if not files:
            self.set_status("没有可回滚的批处理备份", "warn", hold=6)
            messagebox.showinfo("撤销批处理", "这个文件还没有批处理备份。\n（备份只在批处理前自动生成）")
            return
        if not messagebox.askyesno(
                "撤销上次批处理",
                "将丢弃当前未保存的改动，把文件回滚到最近一次备份：\n\n%s\n\n（共 %d 份备份）\n\n继续？"
                % (os.path.basename(files[-1]), len(files))):
            return
        self._busy = True
        self.set_status("正在回滚…", hold=600)
        self.root.update_idletasks()
        try:
            ok, msg = self.bridge.restore_last_backup()
        finally:
            self._busy = False
        if ok:
            self._line_geo = None
            self.current_key = None
            self._refresh_docs()
            self.set_status(msg, "ok", hold=8)
        else:
            self.set_status(msg, "bad", hold=8)
            messagebox.showwarning("撤销失败", msg)

    def _check_activate(self):
        """第二个实例写下的唤醒标记：把自己展开、居中、置顶。"""
        try:
            if not os.path.exists(ACTIVATE_FLAG):
                return
            m = os.path.getmtime(ACTIVATE_FLAG)
            if m <= getattr(self, "_act_seen", 0):
                return
            self._act_seen = m
            if self.collapsed:
                self._toggle_collapse()
            self._center_self()
            self.set_status("助手已在运行 —— 已把它放到屏幕中央", "ok", hold=8)
        except Exception:
            pass

    def _center_self(self):
        """用 Win32 坐标居中（和 DPI 无关，避免 Tk 逻辑坐标和物理像素错位）。"""
        try:
            hwnd = int(self.root.frame(), 16)
            if center_window(hwnd):
                return
        except Exception:
            pass
        try:
            self.root.update_idletasks()
            w = self.root.winfo_width() or W
            h = self.root.winfo_height() or self._full_height()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.root.geometry("%dx%d+%d+%d"
                               % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 2)))
        except Exception:
            pass

    def _cleanup_backup_cache(self):
        """文档保存并关闭后，清掉它的备份缓存。"""
        if not self._backup_track:
            return
        open_paths = set()
        try:
            for i in range(1, int(self.bridge.app.Documents.Count) + 1):
                d = self.bridge.app.Documents.Item(i)
                try:
                    open_paths.add(os.path.normcase(str(d.FullFileName)))
                except Exception:
                    pass
        except Exception:
            return
        for path, t in list(self._backup_track.items()):
            if os.path.normcase(path) in open_paths:
                continue
            try:
                saved = os.path.getmtime(path) > t + 1
            except Exception:
                saved = False
            if saved:
                clear_backups(path)
                log("文档已保存并关闭，清理备份：%s" % path)
            self._backup_track.pop(path, None)

    def _build_linescan(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "线条检测", fg=MUTED, font=FONT_S).pack(side="left")

        # 右侧从右往左：下限 [0.01] · 上限 [0.05] mm
        self.eps_var = tk.StringVar(value="%g" % float(self.cfg.get("eps_mm", 0.005)))
        self.eps_entry = tk.Entry(head, textvariable=self.eps_var, font=FONT_S, width=4,
                                  justify="center", relief="flat", bg=FIELD_BG, fg=TXT,
                                  insertbackground=TXT, highlightthickness=1,
                                  highlightbackground=LINE, highlightcolor=ACCENT)
        self.eps_entry.pack(side="right", ipady=2)
        self._label(head, "下限", fg=MUTED, font=FONT_S).pack(side="right", padx=(0, 2))

        self.thr_var = tk.StringVar(value="%g" % float(self.cfg.get("thr_mm", 0.2)))
        self.thr_entry = tk.Entry(head, textvariable=self.thr_var, font=FONT_S, width=4,
                                  justify="center", relief="flat", bg=FIELD_BG, fg=TXT,
                                  insertbackground=TXT, highlightthickness=1,
                                  highlightbackground=LINE, highlightcolor=ACCENT)
        self._label(head, "mm", fg=MUTED, font=FONT_S).pack(side="right", padx=(6, 2))
        self.thr_entry.pack(side="right", ipady=2)

        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(8, 0))
        self.btn_bound = tk.Label(row, text="边界检测", bg=CHIP_BG, fg=TXT, font=FONT_S,
                                  pady=6, cursor="hand2")
        self.btn_bound.pack(side="left", fill="x", expand=True)
        self.btn_global = tk.Label(row, text="全图闭合检测", bg=CHIP_BG, fg=TXT,
                                   font=FONT_S, pady=6, cursor="hand2")
        self.btn_global.pack(side="left", fill="x", expand=True, padx=(6, 0))
        for w, mode in ((self.btn_bound, "boundary"), (self.btn_global, "global")):
            w.bind("<Button-1>", lambda e, m=mode: self._run_linescan(m))
            w.bind("<Enter>", lambda e, x=w: x.config(bg=CHIP_HOVER))
            w.bind("<Leave>", lambda e, x=w: x.config(bg=CHIP_BG))

        self.ref_lb = tk.Label(parent, text="", bg=BG, fg=ACCENT, font=FONT_S,
                               anchor="w", cursor="hand2")
        self.ref_lb.pack(fill="x", pady=(6, 0))
        self.ref_lb.bind("<Button-1>", lambda e: self._pick_ref_layers())
        self._paint_ref_layers()

    # ---- 参考图层选择 -----------------------------------------

    def _all_layers(self):
        names = self.bridge.layer_names() if self.bridge.connected else []
        return [n for n in names if n not in ("检测标记",)]

    def _ref_layers(self):
        """当前选中的参考图层（与文档里实际存在的图层取交集）。"""
        names = self._all_layers()
        chosen = [n for n in (self.cfg.get("boundary_layers") or []) if n in names]
        if chosen:
            return chosen
        for d in ("边界", "地质界限", "国界"):
            if d in names:
                return [d]
        return []

    def _paint_ref_layers(self):
        try:
            refs = self._ref_layers()
            if not self.bridge.connected:
                txt = "参考图层：—（未连接）"
            elif refs:
                txt = "参考图层：%s ▾" % "、".join(refs if len(refs) <= 3 else refs[:3] + ["…"])
            else:
                txt = "参考图层：未选（点击选择）"
            self.ref_lb.config(text=txt)
        except Exception:
            pass

    def _pick_ref_layers(self):
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
            self._ever_connected = True
            self.current_key = None
            self._refresh_docs()
        names = self._all_layers()
        if not names:
            messagebox.showinfo("参考图层", "当前文档还没有图层")
            return
        cur = set(self._ref_layers())

        top = tk.Toplevel(self.root)
        top.title("选择参考图层")
        top.configure(bg=BG)
        top.attributes("-topmost", True)
        top.geometry("340x470")
        tk.Label(top, text="边界检测以这些图层的线条为参考", bg=BG, fg=TXT,
                 font=FONT_B).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(top, text="勾选任意数量；未勾选则退回「边界」", bg=BG, fg=MUTED,
                 font=FONT_S).pack(anchor="w", padx=14, pady=(0, 8))

        bar = tk.Frame(top, bg=BG)
        bar.pack(side="bottom", fill="x", padx=14, pady=10)

        holder = tk.Frame(top, bg=BG)
        holder.pack(fill="both", expand=True, padx=14)
        canvas = tk.Canvas(holder, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(holder, command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        vars_map = {}
        for n in names:
            v = tk.BooleanVar(value=(n in cur))
            vars_map[n] = v
            cb = tk.Checkbutton(inner, text=n, variable=v, bg=BG, fg=TXT, font=FONT_S,
                                activebackground=BG, activeforeground=TXT,
                                selectcolor=FIELD_BG, anchor="w")
            cb.pack(fill="x", pady=1)

        def sel_all(flag):
            for v in vars_map.values():
                v.set(flag)

        tk.Button(bar, text="全选", command=lambda: sel_all(True)).pack(side="left")
        tk.Button(bar, text="清空", command=lambda: sel_all(False)).pack(side="left", padx=6)

        def ok():
            picked = [n for n, v in vars_map.items() if v.get()]
            self.cfg["boundary_layers"] = picked
            self._save_cfg()
            self._paint_ref_layers()
            self.set_status("参考图层：%s" % ("、".join(picked) if picked else "未选"), hold=5)
            top.destroy()

        tk.Button(bar, text="确定", command=ok).pack(side="right")
        tk.Button(bar, text="取消", command=top.destroy).pack(side="right", padx=6)

    def _eps_mm(self):
        try:
            v = float(self.eps_var.get().strip())
        except Exception:
            v = 0.005
        v = max(0.0, min(1.0, v))
        self.eps_var.set("%g" % v)
        try:
            self.cfg["eps_mm"] = v
            self._save_cfg()
        except Exception:
            pass
        return v

    def _thr_mm(self):
        try:
            v = float(self.thr_var.get().strip())
        except Exception:
            v = 0.2
        v = max(0.005, min(20.0, v))
        self.thr_var.set("%g" % v)
        try:
            self.cfg["thr_mm"] = v
            self._save_cfg()
        except Exception:
            pass
        return v

    def _run_linescan(self, mode):
        if self._busy:
            return
        if linecheck is None:
            messagebox.showerror("缺少模块", "linecheck.py 没有加载成功")
            return
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
            self._ever_connected = True
            self.current_key = None
            self._refresh_docs()
        if not self.bridge.ensure_doc():
            self.set_status("没有打开的文档", "warn")
            return

        thr = self._thr_mm()
        eps = self._eps_mm()
        title = "边界检测" if mode == "boundary" else "全图闭合检测"
        refs = self._ref_layers()
        if mode == "boundary":
            if not refs:
                messagebox.showwarning(title, "请先点「参考图层」选择至少一个图层")
                return
            tip = ("检查所有开放线条的端点，是否贴在「%s」的线条上却没接住\n"
                   "（缝隙 %.3f ~ %.2f mm；小于 %.3f mm 视为已闭合）"
                   % ("、".join(refs), eps, thr, eps))
        else:
            tip = ("检查所有开放线条的端点，是否贴在别的线条上却没接住\n"
                   "（缝隙 %.3f ~ %.2f mm；小于 %.3f mm 视为已闭合）" % (eps, thr, eps))
        if not messagebox.askyesno(title, tip + "\n\n"
                                   "· 只检测有端点的线条，闭合线不参与\n"
                                   "· 首次扫描要复制一份矢量做几何分析，约 1 分钟\n"
                                   "· 不改原图：先标红圈，再逐条确认是否延长\n\n开始？"):
            return

        self._busy = True
        self._cancel_pending()
        dlg = ProgressDialog(self, title, "正在采集线条几何…（首次约 1 分钟）")
        cancelled = False

        def progress(frac, text):
            dlg.set(frac * 0.9, text)
            dlg.check()

        try:
            issues, err = self.bridge.scan_lines(mode, thr, ref_layers=refs,
                                                 progress=progress, eps_mm=eps)
            dlg.set(0.97, "正在比对端点…")
        except Cancelled:
            cancelled = True
            issues, err = None, ""
        finally:
            self._busy = False
            dlg.close()

        if cancelled:
            self.set_status("已取消检测（临时文档已清理，原图未改）", "warn", hold=8)
            return

        if issues is None:
            self.set_status(err or "检测失败", "bad", hold=8)
            messagebox.showwarning(title, err or "检测失败")
            return
        if not issues:
            self.set_status("没有发现未闭合端点（阈值 %.2f mm）" % thr, "ok", hold=8)
            messagebox.showinfo(title, "没有发现未闭合端点。\n阈值 %.2f mm 内都接上了。" % thr)
            return

        dlg2 = ProgressDialog(self, title, "正在标注端点…")
        mark_cancelled = False

        def mprog(frac, text):
            dlg2.set(frac, text)
            dlg2.check()

        self._busy = True
        try:
            n_mark = self.bridge.mark_issues(issues, progress=mprog)
        except Cancelled:
            n_mark = 0
            mark_cancelled = True
        finally:
            self._busy = False
            dlg2.close()

        if mark_cancelled:
            try:
                self.bridge.clear_marks()
            except Exception:
                pass
            self.set_status("已取消标注（检测结果未保留）", "warn", hold=8)
            return
        g = getattr(self.bridge, "_graph", None)
        if g:
            eff = sum(1 for i in issues if i.get("will_gain"))
            self.set_status("%s：%d 处（有效 %d）· 当前可填面 F=%d，修完可到 %d"
                            % (title, len(issues), eff, g["F"], g["F"] + eff),
                            "ok", hold=12)
        else:
            self.set_status("%s：%d 处" % (title, len(issues)), "ok", hold=10)
        self._open_review(issues, mode, thr, n_mark)

    def _apply_all(self, issues, win=None):
        """对给定的检测结果整体执行闭合（先备份，可取消，可一次撤销）。"""
        if self._busy or not issues:
            return
        if not self.bridge.ensure_doc():
            self.set_status("没有打开的文档", "warn")
            return

        gaps = [i for i in issues if i.get("kind") == "gap"]
        ends = [i for i in issues if i.get("kind") != "gap"]
        eff = [i for i in issues if i.get("will_gain")]
        low = [i for i in issues if i.get("kind") == "end" and not i.get("will_gain")]
        par = []
        targets = eff
        if not targets:
            messagebox.showinfo("一键闭合", "没有可自动修复的项（剩下的都是跨片连接）。")
            return
        g = getattr(self.bridge, "_graph", None)
        base_f = g["F"] if g else 0
        if not messagebox.askyesno(
                "一键闭合",
                "将对 %d 处「同一片里断开的线头」执行闭合：\n\n"
                "· 当前可填面 F = %d，每修一处 +1，修完约 %d\n"
                "· 另有 %d 处接上后不会多出面（跨片或与已有连接重复），跳过\n\n"
                "动手前会自动备份，完成后可用底栏「撤销」一次性回退。\n\n继续？"
                % (len(targets), base_f, base_f + len(targets), len(low))):
            return

        self._backup_for_batch("一键闭合")
        dlg = ProgressDialog(self, "一键闭合", "准备中…")
        ok = fail = 0
        errs = []
        cancelled = False
        self._busy = True
        try:
            try:
                self.bridge.doc.BeginCommandGroup("一键闭合")
            except Exception:
                pass
            n = len(targets)
            for i, it in enumerate(targets):
                dlg.set(i / float(n), "处理 %d/%d" % (i + 1, n))
                dlg.check()
                if it.get("kind") == "gap":
                    good, msg = self.bridge.fix_gap(it)
                else:
                    good, msg = self.bridge.extend_issue(it)
                if good:
                    ok += 1
                else:
                    fail += 1
                    if len(errs) < 5:
                        errs.append("%s：%s" % (it.get("layer", ""), msg))
            try:
                self.bridge.doc.EndCommandGroup()
            except Exception:
                pass
        except Cancelled:
            cancelled = True
            try:
                self.bridge.doc.EndCommandGroup()
            except Exception:
                pass
        finally:
            self._busy = False
            dlg.close()

        self._line_geo = None
        try:
            self.bridge.clear_marks()
        except Exception:
            pass
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass

        head = "已取消，" if cancelled else ""
        self.set_status("%s闭合 %d 处，失败 %d 处" % (head, ok, fail),
                        "ok" if not fail else "warn", hold=10)
        detail = "%s一键闭合：成功 %d，失败 %d" % (head, ok, fail)
        if low:
            detail += "\n跳过跨片连接 %d 处（不增加可填面）" % len(low)
        if errs:
            detail += "\n\n失败示例：\n" + "\n".join(errs)
        detail += "\n\n可用底栏「撤销」一次性回退。"
        messagebox.showinfo("一键闭合", detail)

    def _open_review(self, issues, mode, thr, n_mark):
        title = "边界检测" if mode == "boundary" else "全图闭合检测"
        win = tk.Toplevel(self.root)
        win.title("%s · %d 处" % (title, len(issues)))
        win.configure(bg=BG)
        win.attributes("-topmost", True)
        sw = win.winfo_screenwidth()
        win.geometry("560x565+%d+%d" % (max(0, sw - 640), 70))

        head = tk.Frame(win, bg=BG)
        head.pack(fill="x", padx=14, pady=(12, 2))
        tk.Label(head, text="%s · 阈值 %.2f mm" % (title, thr), bg=BG, fg=TXT,
                 font=FONT_B).pack(side="left")
        tk.Label(head, text="已标 %d 个红圈" % n_mark, bg=BG, fg=MUTED,
                 font=FONT_S).pack(side="right")

        tk.Label(win, text="选一行 → 定位；确认后用「延长到交点」把这一端补到线上。",
                 bg=BG, fg=MUTED, font=FONT_S, anchor="w").pack(fill="x", padx=14, pady=(2, 6))

        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=14)
        sb = tk.Scrollbar(body)
        sb.pack(side="right", fill="y")
        lb = tk.Listbox(body, font=FONT_S, bg=FIELD_BG, fg=TXT, relief="flat",
                        highlightthickness=1, highlightbackground=LINE,
                        selectbackground=SEL_BG, selectforeground=SEL_FG,
                        yscrollcommand=sb.set, activestyle="none")
        lb.pack(side="left", fill="both", expand=True)
        sb.config(command=lb.yview)

        items = list(issues)
        cur = {"i": 0, "backed": False}

        def fmt(it):
            if it.get("kind") == "end":
                mark = "→" if it.get("same_comp") else "·"
                return "%s%02d  %-8s %-3s %6.3fmm  (%.1f, %.1f)" % (
                    mark, it["no"], it["layer"], "尾端" if it["end"] else "首端",
                    it["dist_mm"], it["pt"][0], it["pt"][1])
            if it.get("kind") == "gap":
                if it.get("parallel"):
                    mark, tail = "∥", "疑似重复"
                elif it.get("low_risk"):
                    mark, tail = "·", "低风险"
                else:
                    mark, tail = "→", "两线之间"
            else:
                mark = (" " if it.get("mode") == "ray"
                        else ("⊥" if it.get("mode") == "perp" else "×"))
                tail = "尾端" if it["end"] else "首端"
            return "%s%02d  %-8s %-6s %5.3fmm  (%.1f, %.1f)" % (
                mark, it["no"], it["layer"], tail,
                it["dist_mm"], it["pt"][0], it["pt"][1])

        def refresh():
            lb.delete(0, "end")
            for it in items:
                lb.insert("end", fmt(it))
            if items:
                idx = min(cur["i"], len(items) - 1)
                lb.selection_clear(0, "end")
                lb.selection_set(idx)
                lb.see(idx)
            try:
                btn_all.config(text="一键闭合全部（剩余 %d 处）" % len(items),
                               state="normal" if items else "disabled")
            except Exception:
                pass

        def current():
            if not items:
                return None
            sel = lb.curselection()
            cur["i"] = sel[0] if sel else 0
            return items[cur["i"]]

        def on_select(_e=None):
            it = current()
            if it:
                self.bridge.select_shape(it)

        lb.bind("<<ListboxSelect>>", on_select)

        def do_zoom():
            it = current()
            if it:
                self._zoom_to(it)

        def do_extend():
            it = current()
            if not it:
                return
            if it.get("kind") == "gap":
                ok, msg = self.bridge.fix_gap(it)
                self.set_status(msg, "ok" if ok else "bad", hold=6)
                if ok:
                    items.pop(cur["i"])
                    refresh()
                    on_select()
                return
            if not it.get("extend_to"):
                messagebox.showwarning(
                    "无法自动延长",
                    (it.get("reason") or "射线方向打不到参考线") +
                    "\n\n这一处需要手动补，或用「跳过」略过。")
                return
            if not cur.get("backed"):
                self._backup_for_batch("线条延长")
                cur["backed"] = True
            ok, msg = self.bridge.extend_issue(it)
            self.set_status(msg, "ok" if ok else "bad", hold=6)
            if ok:
                items.pop(cur["i"])
                refresh()
                on_select()

        def do_skip():
            if not current():
                return
            items.pop(cur["i"])
            refresh()
            on_select()

        def do_close():
            self.bridge.clear_marks()
            self.set_status("检测结束，标记已清除", hold=5)
            win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=14, pady=(10, 12))

        def mk(text, cmd, bg, fg):
            w = tk.Label(btns, text=text, bg=bg, fg=fg, font=FONT_S, padx=10, pady=6,
                         cursor="hand2")
            w.pack(side="left", padx=(0, 6))
            w.bind("<Button-1>", lambda e: cmd())
            return w

        mk("定位 / 放大", do_zoom, CHIP_BG, TXT)
        mk("延长到交点", do_extend, ACCENT, "#FFFFFF")
        mk("跳过", do_skip, CHIP_BG, TXT)
        close_btn = tk.Label(btns, text="关闭", bg=CHIP_BG, fg=TXT, font=FONT_S,
                             padx=10, pady=6, cursor="hand2")
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: do_close())

        allrow = tk.Frame(win, bg=BG)
        allrow.pack(fill="x", padx=14, pady=(8, 12))
        btn_all = tk.Label(allrow, text="一键闭合全部", bg=ACCENT, fg="#FFFFFF",
                           font=FONT_B, pady=8, cursor="hand2")
        btn_all.pack(fill="x")
        btn_all.bind("<Button-1>", lambda e: self._apply_all(list(items), win))
        btn_all.bind("<Enter>", lambda e: btn_all.config(bg=ACCENT_HOVER))
        btn_all.bind("<Leave>", lambda e: btn_all.config(bg=ACCENT))

        win.protocol("WM_DELETE_WINDOW", do_close)
        refresh()
        on_select()

    def _cdr_window(self):
        try:
            name = str(self.bridge.doc.Name)
        except Exception:
            return 0
        found = []

        def cb(h, _):
            try:
                t = win32gui.GetWindowText(h)
            except Exception:
                return
            if name in t and win32gui.IsWindowVisible(h):
                found.append(h)
        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            return 0
        return found[0] if found else 0

    def _zoom_to(self, it):
        if not self.bridge.select_shape(it):
            self.set_status("定位失败：找不到这条线", "bad", hold=5)
            return
        try:
            hwnd = self._cdr_window()
            if hwnd:
                win32gui.ShowWindow(hwnd, 3)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.2)
                win32api.keybd_event(0x10, 0, 0, 0)
                win32api.keybd_event(0x71, 0, 0, 0)
                time.sleep(0.08)
                win32api.keybd_event(0x71, 0, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as exc:
            log("zoom 失败: %r" % (exc,))
        self.set_status("已定位：%s 第 %d 处" % (it["layer"], it["no"]), hold=4)

    def _build_bar(self):
        bar = tk.Frame(self.shell, bg=BG, height=BAR_H)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self.dot = tk.Canvas(bar, width=10, height=10, bg=BG, highlightthickness=0)
        self.dot.pack(side="left", padx=(12, 6))
        self.dot_id = self.dot.create_oval(1, 1, 9, 9, fill=BAD, outline="")

        self.title = tk.Label(bar, text="CorelDRAW 助手", bg=BG, fg=TXT, font=FONT_B)
        self.title.pack(side="left")

        self.ver_lb = tk.Label(bar, text="v" + APP_VERSION, bg=BG, fg=MUTED,
                               font=FONT_S, cursor="hand2")
        self.ver_lb.pack(side="left", padx=(6, 0))
        self.ver_lb.bind("<Button-1>", lambda e: self._check_update(manual=True))
        self.ver_lb.bind("<Enter>", lambda e: self.ver_lb.config(fg=ACCENT))
        self.ver_lb.bind("<Leave>", lambda e: self.ver_lb.config(fg=MUTED))

        self.btn_close = tk.Label(bar, text="✕", bg=BG, fg=MUTED, font=FONT, padx=8, cursor="hand2")
        self.btn_close.pack(side="right", padx=(0, 4))
        self.btn_close.bind("<Button-1>", lambda e: self._quit())
        self.btn_close.bind("<Enter>", lambda e: self.btn_close.config(fg=BAD))
        self.btn_close.bind("<Leave>", lambda e: self.btn_close.config(fg=MUTED))

        self.btn_min = tk.Label(bar, text="—", bg=BG, fg=MUTED, font=FONT, padx=8, cursor="hand2")
        self.btn_min.pack(side="right")
        self.btn_min.bind("<Button-1>", lambda e: self._toggle_collapse())
        self.btn_min.bind("<Enter>", lambda e: self.btn_min.config(fg=TXT))
        self.btn_min.bind("<Leave>", lambda e: self.btn_min.config(fg=MUTED))

        self.refresh_btn = tk.Label(bar, text="⟳", bg=BG, fg=MUTED, font=FONT, padx=6, cursor="hand2")
        self.refresh_btn.pack(side="right")
        self.refresh_btn.bind("<Button-1>", lambda e: self._connect(manual=True))
        self.refresh_btn.bind("<Enter>", lambda e: self.refresh_btn.config(fg=TXT))
        self.refresh_btn.bind("<Leave>", lambda e: self.refresh_btn.config(fg=MUTED))

        for w in (bar, self.title):
            w.bind("<Button-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag_move)
            w.bind("<ButtonRelease-1>", lambda e: self._save_cfg())

    def _build_docbar(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="x", padx=14, pady=(0, 6))

        self.doc_canvas = tk.Canvas(wrap, bg=BG, height=DOCBAR_H, bd=0,
                                    highlightthickness=0, xscrollincrement=1)
        self.doc_canvas.pack(fill="x")
        self.doc_canvas.bind("<MouseWheel>", self._on_doc_wheel)
        self._render_docbar()

    def _render_docbar(self):
        c = self.doc_canvas
        c.delete("all")
        self._doc_items = []

        if not self.doc_list:
            c.create_text(2, DOCBAR_H / 2.0, text="未连接 CorelDRAW", anchor="w",
                          fill=MUTED, font=FONT_S)
            c.configure(scrollregion=(0, 0, 160, DOCBAR_H))
            return

        x = 0
        y0, y1 = 1, DOCBAR_H - 1
        for n, (idx, name) in enumerate(self.doc_list):
            active = (name == self.current_key)
            label = name if len(name) <= 12 else name[:11] + "…"
            w = max(66, min(150, len(label) * 8 + 24))
            tag = "doc%d" % n
            c.create_rectangle(x, y0, x + w - 4, y1, outline="",
                               fill=SEL_BG if active else TAG_BG, tags=(tag,))
            c.create_text(x + (w - 4) / 2.0, DOCBAR_H / 2.0, text=label,
                          fill=SEL_FG if active else TXT, font=FONT_S, tags=(tag,))
            c.tag_bind(tag, "<Button-1>", lambda e, k=name, i=idx: self._switch_doc(k, i))
            x += w

        c.configure(scrollregion=(0, 0, max(x, 120), DOCBAR_H))

    def _on_doc_wheel(self, e):
        self.doc_canvas.xview_scroll(-30 if e.delta > 0 else 30, "units")
        return "break"

    def _sep(self, parent):
        tk.Frame(parent, bg=LINE, height=1).pack(fill="x", pady=8)

    def _build_width(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "线宽", fg=MUTED, font=FONT_S).pack(side="left")

        self.width_var = tk.StringVar(value=self._fmt_mm(self.state["width_mm"]))
        self.entry = tk.Entry(head, textvariable=self.width_var, font=FONT_NUM,
                              width=6, justify="center", relief="flat",
                              bg=FIELD_BG, fg=TXT, insertbackground=TXT,
                              highlightthickness=1, highlightbackground=LINE,
                              highlightcolor=ACCENT)
        self._label(head, "mm", fg=MUTED, font=FONT_S).pack(side="right", padx=(5, 0))
        self.entry.pack(side="right", ipady=3)

        self.entry.bind("<Return>", lambda e: self._commit_width())
        self.entry.bind("<FocusOut>", lambda e: self._commit_width())

        chips = tk.Frame(parent, bg=BG)
        chips.pack(fill="x", pady=(7, 0))
        for text, val in PRESET_WIDTHS:
            self._chip(chips, text, lambda v=val: self._set_width(v)).pack(side="left", padx=(0, 4))

    def _build_color(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "轮廓颜色", fg=MUTED, font=FONT_S).pack(side="left")

        self._chip(head, "…", self._pick_color, width=2).pack(side="right")
        self.cur_color = tk.Frame(head, bg=rgb_hex(self.state["outline_rgb"]),
                                  width=16, height=16,
                                  highlightthickness=1, highlightbackground="#C9C7BF")
        self.cur_color.pack(side="right", padx=(0, 6))
        self.cur_color.pack_propagate(False)

        self.pal_hint = self._label(head, "", fg=MUTED, font=FONT_S)
        self.pal_hint.pack(side="right", padx=(0, 8))

        holder = tk.Frame(parent, bg=BG)
        holder.pack(fill="x", pady=(8, 0))

        self.pal_canvas = tk.Canvas(holder, bg=BG, height=PAL_H, bd=0,
                                    highlightthickness=0,
                                    yscrollincrement=PAL_CELL)
        self.pal_canvas.pack(fill="x")
        self.pal_canvas.bind("<MouseWheel>", self._on_pal_wheel)

        self.palette_colors = []
        self._pal_items = {}
        self._pal_mark = None
        self._render_palette()

    def _render_palette(self):
        c = self.pal_canvas
        c.delete("all")
        self._pal_items = {}
        self._pal_mark = None

        cols = self.palette_colors or FALLBACK_COLORS
        for i, rgb in enumerate(cols):
            r, k = divmod(i, PAL_COLS)
            x = k * PAL_CELL + 3
            y = r * PAL_CELL + 3
            item = c.create_rectangle(x, y, x + PAL_SW, y + PAL_SW,
                                      fill=rgb_hex(rgb), outline="#D3D1C8", width=1)
            c.tag_bind(item, "<Button-1>", lambda e, col=tuple(rgb): self._set_color(col))
            self._pal_items.setdefault(tuple(rgb), (x, y))

        rows = max(PAL_ROWS, (len(cols) + PAL_COLS - 1) // PAL_COLS)
        c.configure(scrollregion=(0, 0, PAL_COLS * PAL_CELL, rows * PAL_CELL))
        c.yview_moveto(0)
        self._mark_palette()

        if hasattr(self, "pal_hint"):
            if rows > PAL_ROWS:
                self.pal_hint.config(text="%d 色 · 滚轮翻" % len(cols))
            else:
                self.pal_hint.config(text="%d 色" % len(cols))

    def _mark_palette(self):
        try:
            if self._pal_mark:
                self.pal_canvas.delete(self._pal_mark)
        except Exception:
            pass
        self._pal_mark = None
        pos = getattr(self, "_pal_items", {}).get(tuple(self.state["outline_rgb"]))
        if pos:
            x, y = pos
            self._pal_mark = self.pal_canvas.create_rectangle(
                x - 2, y - 2, x + PAL_SW + 2, y + PAL_SW + 2,
                outline=ACCENT, width=2)

    def _on_pal_wheel(self, e):
        self.pal_canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")
        return "break"

    def _build_fill(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "填充", fg=MUTED, font=FONT_S).pack(side="left")

        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(9, 0))

        self.fill_btns = {}
        for key, text in (("none", "无填充"), ("outline", "同轮廓色"), ("custom", "自定色")):
            lb = tk.Label(row, text=text, font=FONT_S, padx=8, pady=5, cursor="hand2")
            lb.pack(side="left", padx=(0, 5))
            lb.bind("<Button-1>", lambda e, k=key: self._set_fill(k))
            self.fill_btns[key] = lb

        self.fill_swatch = tk.Frame(row, bg=rgb_hex(self.state["fill_rgb"]),
                                    width=18, height=18,
                                    highlightthickness=1, highlightbackground="#D3D1C8",
                                    cursor="hand2")
        self.fill_swatch.pack(side="left", padx=(3, 0))
        self.fill_swatch.pack_propagate(False)
        self.fill_swatch.bind("<Button-1>", lambda e: self._pick_fill())
        self._paint_fill()

    def _build_targets(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "应用对象", fg=MUTED, font=FONT_S).pack(side="left")

        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(8, 0))

        self.kind_btns = {}
        for key, text in (("graphic", "图形 / 曲线"), ("artistic", "艺术笔")):
            lb = tk.Label(row, text=text, font=FONT_S, padx=9, pady=5, cursor="hand2")
            lb.pack(side="left", padx=(0, 6))
            lb.bind("<Button-1>", lambda e, k=key: self._toggle_kind(k))
            self.kind_btns[key] = lb
        self._paint_kinds()

        live_row = tk.Frame(parent, bg=BG)
        live_row.pack(fill="x", pady=(9, 0))
        self.live_lb = tk.Label(live_row, text="", font=FONT_S, bg=BG, cursor="hand2", pady=2)
        self.live_lb.pack(side="left")
        self.live_lb.bind("<Button-1>", lambda e: self._toggle_live())

        self.auto_lb = tk.Label(live_row, text="", font=FONT_S, bg=BG, cursor="hand2", pady=2)
        self.auto_lb.pack(side="right")
        self.auto_lb.bind("<Button-1>", lambda e: self._toggle_autostart())
        self._paint_live()

    def _build_batch(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x")
        self._label(head, "批量处理", fg=MUTED, font=FONT_S).pack(side="left")
        self._label(head, "全文档 · 全部图层", fg=MUTED, font=FONT_S).pack(side="right")

        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(8, 0))
        self.btn_group = tk.Label(row, text="按线型分组", bg=ACCENT_SOFT, fg=ACCENT,
                                  font=FONT_S, pady=7, cursor="hand2")
        self.btn_group.pack(side="left", fill="x", expand=True)
        self.btn_color = tk.Label(row, text="同色归并", bg=ACCENT_SOFT, fg=ACCENT,
                                  font=FONT_S, pady=7, cursor="hand2")
        self.btn_color.pack(side="left", fill="x", expand=True, padx=(6, 0))

        self.btn_undo = tk.Label(parent, text="回滚到上次批处理前（用备份）", bg=CHIP_BG,
                                 fg=MUTED, font=FONT_S, pady=5, cursor="hand2")
        self.btn_undo.pack(fill="x", pady=(6, 0))

        for w, cmd in ((self.btn_group, self._group_by_linetype),
                       (self.btn_color, self._run_color_merge),
                       (self.btn_undo, self._undo_batch)):
            w.bind("<Button-1>", lambda e, c=cmd: c())
        for w in (self.btn_group, self.btn_color):
            w.bind("<Enter>", lambda e, x=w: x.config(bg=ACCENT_SOFT_HOVER))
            w.bind("<Leave>", lambda e, x=w: x.config(bg=ACCENT_SOFT))
        self.btn_undo.bind("<Enter>", lambda e: self.btn_undo.config(bg=CHIP_HOVER))
        self.btn_undo.bind("<Leave>", lambda e: self.btn_undo.config(bg=CHIP_BG))

    # ---- 交互 -------------------------------------------------

    def _fmt_mm(self, v):
        if v is None:
            return ""
        if float(v) == 0:
            return "0"
        return ("%.3f" % float(v)).rstrip("0").rstrip(".")

    def _drag_start(self, e):
        self._drag = (e.x_root - self.root.winfo_x(), e.y_root - self.root.winfo_y())

    def _drag_move(self, e):
        if self._drag:
            self.root.geometry("+%d+%d" % (e.x_root - self._drag[0], e.y_root - self._drag[1]))

    def _toggle_collapse(self):
        self.collapsed = not self.collapsed
        try:
            self.btn_min.config(text="▾" if self.collapsed else "—")
        except Exception:
            pass
        if self.collapsed:
            self.body.pack_forget()
            self.quickbar.pack_forget()
            self.verbar.pack_forget()
            self.root.geometry("%dx%d" % (W, BAR_H + 2))
        else:
            self.quickbar.pack(side="bottom", fill="x")
            self.verbar.pack(side="bottom", fill="x")
            self.body.pack(fill="both", expand=True)
            self.root.geometry("%dx%d" % (W, self._full_height()))

    def _quit(self):
        self._save_cfg()
        try:
            if getattr(self, "tray", None) is not None:
                self.tray.stop()
        except Exception:
            pass
        self.root.destroy()

    def set_status(self, text, level="info", hold=0):
        """hold>0 时，这几秒内不让人手/心跳的常规提示把它顶掉。"""
        if hold > 0:
            self._status_hold = time.time() + hold
        self.status.config(text=text,
                           fg={"info": MUTED, "ok": OK, "warn": WARN, "bad": BAD}.get(level, MUTED))

    def _status_free(self):
        return time.time() >= getattr(self, "_status_hold", 0)

    # ---- 自动更新：启动静默检查，只有线上有新版本才打扰 ----------

    def _check_update(self, manual=False):
        if self._update_checked and not manual:
            return
        self._update_checked = True
        self._update_manual = manual
        self._update_done = False
        if manual:
            self.set_status("正在检查更新…", hold=2)
            self._set_ver_state("正在检查更新…", "info")
        threading.Thread(target=self._fetch_update_bg, daemon=True).start()

    def _fetch_update_bg(self):
        """后台线程：只写属性，绝不碰 Tk 控件（Tk 不是线程安全的）。"""
        try:
            info = fetch_update_info()
        except Exception as exc:
            info = {"error": "%s: %s" % (type(exc).__name__, exc)}
        self._update_info = info
        self._update_done = True

    def _poll_update(self):
        """主线程里收后台结果：没有更新就什么都不做。"""
        if not self._update_done:
            return
        self._update_done = False
        info = self._update_info or {}
        manual = self._update_manual
        if info.get("error"):
            log("检查更新失败: %s" % info["error"])
            self._update_latest = None
            self._set_ver_state("检查更新失败 · 点 ⟳ 重试", "warn")
            if manual:
                self.set_status("检查更新失败：%s" % str(info["error"])[:36], "warn", hold=6)
            return
        latest = info.get("version") or ""
        if not version_tuple(latest):
            log("检查更新：线上没有可识别的版本号（%s）" % info.get("source", "?"))
            self._update_latest = None
            self._set_ver_state("线上没有可识别的版本号", "warn")
            if manual:
                self.set_status("线上没有可识别的版本号", "warn", hold=6)
            return
        if version_tuple(latest) <= version_tuple(APP_VERSION):
            log("检查更新：已是最新 v%s（线上 v%s，来源 %s）"
                % (APP_VERSION, latest, info.get("source", "?")))
            self._update_latest = None
            self._set_ver_state("当前已是最新版本 v%s" % APP_VERSION, "ok")
            if manual:
                self.set_status("已是最新版本 v%s" % APP_VERSION, "ok", hold=5)
            return
        self._update_latest = info
        self._set_ver_state("检测到新版本 v%s · 点击查看" % latest, "new", has_new=True)
        self._prompt_update(info)

    def _prompt_update(self, info):
        latest = info.get("version")
        notes = (info.get("notes") or "").strip()
        source = info.get("source") or ""
        can_update = exe_path() is not None
        lines = ["发现新版本 v%s（当前 v%s）" % (latest, APP_VERSION)]
        if source:
            lines.append("版本来源：%s" % source)
        if notes:
            lines.append("更新内容：\n%s" % notes[:400])
        lines.append("点「立即更新」自动下载替换并重启。" if can_update
                     else "点「立即更新」打开下载页。")
        self.set_status("发现新版本 v%s" % latest, "ok", hold=10)
        log("发现新版本 v%s（来源 %s）" % (latest, source))
        if not ask_centered(self.root, "发现新版本", lines, ok_text="立即更新"):
            return

        url = pick_download_url(info)
        log("更新下载地址: %s" % url)
        if not can_update:                 # 源码模式没法替换自己
            try:
                webbrowser.open(url)
            except Exception as exc:
                log("打开下载地址失败 %r: %r" % (url, exc))
            return
        self._start_download(url)

    # ---- 一键替换 -------------------------------------------------

    def _start_download(self, url):
        dest = exe_path() + ".new"
        self._dl = {"url": url, "dest": dest, "got": 0, "total": 0, "done": None}
        self.set_status("正在下载更新…", hold=3600)
        threading.Thread(target=self._download_bg, args=(url, dest), daemon=True).start()
        self.root.after(200, self._poll_download)

    def _download_bg(self, url, dest):
        """后台线程：只写字典，不碰 Tk。"""
        st = self._dl
        try:
            download_file(url, dest, lambda g, t: st.update(got=g, total=t))
            if not looks_like_exe(dest):
                raise IOError("下载到的不是有效的程序文件")
            st["done"] = "ok"
        except Exception as exc:
            st["done"] = "err"
            st["error"] = "%s: %s" % (type(exc).__name__, exc)

    def _poll_download(self):
        st = getattr(self, "_dl", None)
        if not st:
            return
        if st.get("done") == "ok":
            self.set_status("下载完成，正在替换…", "ok", hold=3600)
            self.root.update_idletasks()
            if install_update(st["dest"]):
                try:
                    self._save_cfg()
                except Exception:
                    pass
                self.root.destroy()        # 退出，让新进程接手
                return
            messagebox.showwarning("更新失败", "替换程序文件失败，已改为打开下载页。")
            self._open_fallback(st)
            return
        if st.get("done") == "err":
            log("下载更新失败: %s" % st.get("error"))
            messagebox.showerror("更新失败", "下载更新失败：\n%s\n\n已改为打开下载页。" % st.get("error"))
            self._open_fallback(st)
            return

        got, total = st.get("got", 0), st.get("total", 0)
        if total:
            self.set_status("正在下载更新 %d%%（%.1f / %.1f MB）"
                            % (got * 100 // total, got / 1048576.0, total / 1048576.0), hold=3600)
        else:
            self.set_status("正在下载更新 %.1f MB…" % (got / 1048576.0), hold=3600)
        self.root.after(300, self._poll_download)

    def _clean_old_once(self):
        """更新后旧版改名成 .old，等旧进程真正退干净再删（最多试 12 次）。"""
        exe = exe_path()
        if not exe or self._old_clean_tries >= 12:
            return
        p = exe + ".old"
        if not os.path.exists(p):
            return
        self._old_clean_tries += 1
        try:
            os.remove(p)
            log("清理残留旧版成功")
        except Exception:
            pass

    def _open_fallback(self, st):
        self.set_status("更新失败，已打开下载页", "bad", hold=8)
        try:
            webbrowser.open(st.get("url") or UPDATE_WAN_URL)
        except Exception as exc:
            log("打开下载页失败: %r" % (exc,))
        try:
            if st.get("dest") and os.path.exists(st["dest"]):
                os.remove(st["dest"])
        except Exception:
            pass

    def _connect(self, manual=False):
        self.set_status("正在连接…")
        self.root.update_idletasks()
        ok, msg = self.bridge.attach()
        self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
        if not ok:
            self.set_status(msg, "bad")
            return

        self._ever_connected = True
        if self.collapsed and self._auto_collapsed:
            self._toggle_collapse()
            self._auto_collapsed = False

        self.current_key = None
        self.pal_cache.clear()
        self.doc_list = []
        self._render_docbar()
        self._refresh_docs()
        self._paint_ref_layers()
        if not self.doc_list:
            self.set_status(msg, "warn")

    def _load_palette(self):
        try:
            cols = self.bridge.read_palette()
        except Exception as exc:
            log("load_palette 失败: %r" % (exc,))
            cols = []
        if cols:
            self.palette_colors = cols
            self._render_palette()

    def _tick(self):
        """心跳：跟随 CorelDRAW 启停，并刷新打开的文档列表。"""
        if getattr(self, "_busy", False):
            self.root.after(1500, self._tick)
            return
        try:
            self._poll_update()
            self._clean_old_once()
            self._cleanup_backup_cache()
            self._check_activate()
            if not self.bridge.connected:
                running = CdrBridge.cdr_running()
                if running:
                    self._connect()
                elif running is False and self._ever_connected and not self.collapsed:
                    self._toggle_collapse()
                    self._auto_collapsed = True
            else:
                self._refresh_docs()
        except Exception as exc:
            log("tick: %r" % (exc,))
        self.root.after(2500 if not self.bridge.connected else 3000, self._tick)

    def _refresh_docs(self):
        docs = self.bridge.list_docs()
        old = [n for _, n in self.doc_list]
        new = [n for _, n in docs]
        self.doc_list = docs

        if not docs:
            if old:
                self._render_docbar()
            self.current_key = None
            self.bridge.doc = None
            if self._status_free():
                self.set_status("已连接 · 无打开文档", "warn")
            return

        if new != old:
            self._render_docbar()

        if self.current_key not in new:
            self.current_key = new[0]
            self.bridge.use_doc(docs[0][0], activate=False)
            self._load_current_doc()

    def _switch_doc(self, key, index):
        if key == self.current_key:
            return

        # Activate() 会泵消息，把挂起的延迟写入提前触发，必须先掐掉
        self._switching = True
        self._cancel_pending()
        try:
            if self.current_key:
                self.doc_states[self.current_key] = self._clone_state(self.state)
            if not self.bridge.use_doc(index, activate=True):
                self.set_status("切换文档失败", "bad")
                return
            self.current_key = key
            self._render_docbar()
            self._load_current_doc()
            self._line_geo = None
            self._paint_ref_layers()
        finally:
            self._switching = False

    @staticmethod
    def _clone_state(st):
        return {
            "width_mm": st["width_mm"],
            "outline_rgb": tuple(st["outline_rgb"]),
            "fill_mode": st["fill_mode"],
            "fill_rgb": tuple(st["fill_rgb"]),
            "kinds": dict(st["kinds"]),
            "live": st["live"],
            "line_caps": st.get("line_caps", 0),
            "line_join": st.get("line_join", 0),
            "justify": st.get("justify", 0),
            "scale_with": st.get("scale_with", False),
            "behind_fill": st.get("behind_fill", False),
        }

    def _load_current_doc(self):
        key = self.current_key
        pal = self.pal_cache.get(key)
        if not pal:
            try:
                pal = self.bridge.read_palette()
            except Exception as exc:
                log("read_palette: %r" % (exc,))
                pal = []
            if pal:
                self.pal_cache[key] = pal
        self.palette_colors = pal or []
        self._render_palette()

        data = self.bridge.read_current()
        if data:
            self._apply_data_to_state(data)

        self._paint_all()
        self.set_status("%s · 色板 %d 色" % (str(key)[:14], len(self.palette_colors)), "ok")

    def _apply_data_to_state(self, data):
        if data.get("width_mm") is not None:
            self.state["width_mm"] = data["width_mm"]
        if data.get("outline_rgb"):
            self.state["outline_rgb"] = data["outline_rgb"]
        for key in ("line_caps", "line_join", "justify", "scale_with", "behind_fill"):
            if data.get(key) is not None:
                self.state[key] = data[key]
        if data.get("fill_mode"):
            self.state["fill_mode"] = data["fill_mode"]
            if data.get("fill_rgb"):
                self.state["fill_rgb"] = data["fill_rgb"]

    def _paint_all(self):
        self.width_var.set(self._fmt_mm(self.state["width_mm"]))
        self.cur_color.config(bg=rgb_hex(self.state["outline_rgb"]))
        self.fill_swatch.config(bg=rgb_hex(self.state["fill_rgb"]))
        self._paint_fill()
        self._paint_kinds()
        self._paint_live()
        if hasattr(self, "adv_lb"):
            self._paint_advanced()
        self._mark_palette()

    def _commit_width(self):
        raw = self.width_var.get().strip().replace(",", ".")
        try:
            v = float(raw) if raw else 0.0
        except ValueError:
            self.width_var.set(self._fmt_mm(self.state["width_mm"]))
            self.set_status("线宽需为数字", "warn")
            return
        self.state["width_mm"] = max(0.0, min(v, 100.0))
        self.width_var.set(self._fmt_mm(self.state["width_mm"]))
        self._live_apply()

    def _set_width(self, v):
        self.state["width_mm"] = v
        self.width_var.set(self._fmt_mm(v))
        self._live_apply()

    def _set_color(self, rgb):
        self.state["outline_rgb"] = tuple(rgb)
        self.cur_color.config(bg=rgb_hex(rgb))
        self._mark_palette()
        self._live_apply()

    def _pick_color(self):
        res = colorchooser.askcolor(color=rgb_hex(self.state["outline_rgb"]), title="选择轮廓颜色")
        if res and res[0]:
            self._set_color(tuple(int(c) for c in res[0]))

    def _set_fill(self, key):
        self.state["fill_mode"] = key
        self._paint_fill()
        if key == "custom":
            self._pick_fill()
        else:
            self._live_apply()

    def _pick_fill(self):
        res = colorchooser.askcolor(color=rgb_hex(self.state["fill_rgb"]), title="选择填充颜色")
        if res and res[0]:
            self.state["fill_rgb"] = tuple(int(c) for c in res[0])
            self.fill_swatch.config(bg=rgb_hex(self.state["fill_rgb"]))
            self.state["fill_mode"] = "custom"
            self._paint_fill()
            self._live_apply()

    def _paint_fill(self):
        for key, lb in self.fill_btns.items():
            if self.state["fill_mode"] == key:
                lb.config(bg=SEL_BG, fg=SEL_FG)
            else:
                lb.config(bg=CHIP_BG, fg=TXT)

    def _toggle_kind(self, key):
        self.state["kinds"][key] = not self.state["kinds"][key]
        self._paint_kinds()
        self._live_apply()

    def _paint_kinds(self):
        for key, lb in self.kind_btns.items():
            if self.state["kinds"].get(key):
                lb.config(bg=SEL_BG, fg=SEL_FG)
            else:
                lb.config(bg=CHIP_BG, fg=TXT)

    def _toggle_live(self):
        self.state["live"] = not self.state["live"]
        self._paint_live()

    def _paint_live(self):
        if self.state["live"]:
            self.live_lb.config(text="✓ 改动立即生效", fg=ACCENT)
        else:
            self.live_lb.config(text="○ 改动需点按钮", fg=MUTED)
        if hasattr(self, "auto_lb"):
            if autostart_enabled():
                self.auto_lb.config(text="✓ 开机自启", fg=ACCENT)
            else:
                self.auto_lb.config(text="○ 开机自启", fg=MUTED)

    def _toggle_autostart(self):
        on = autostart_enabled()
        ok, err = set_autostart(not on)
        if not ok:
            self.set_status(err, "warn")
        else:
            self.set_status("已设为开机自启" if not on else "已取消开机自启",
                            "ok" if not on else "info")
        self._paint_live()

    def _live_apply(self):
        self._save_cfg()
        if not self.state["live"]:
            return
        self._cancel_pending()
        self._pending = self.root.after(120, lambda: self._apply(explicit=False))

    def _cancel_pending(self):
        if self._pending:
            try:
                self.root.after_cancel(self._pending)
            except Exception:
                pass
            self._pending = None

    def _apply(self, explicit=False):
        if getattr(self, "_switching", False):
            return
        self._cancel_pending()
        if not self.bridge.connected:
            ok, msg = self.bridge.attach()
            self.dot.itemconfig(self.dot_id, fill=OK if ok else BAD)
            if not ok:
                self.set_status(msg, "bad")
                return
            self._ever_connected = True
            self.current_key = None
            self._refresh_docs()

        if self.current_key and self.bridge.doc_name != self.current_key:
            for idx, name in self.doc_list:
                if name == self.current_key:
                    self.bridge.use_doc(idx, activate=False)
                    break

        done, msg = self.bridge.apply(self.state)
        self.set_status(msg, "ok" if done else "warn")
        self.dot.itemconfig(self.dot_id, fill=OK if done else WARN)

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------- 入口

def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    cleanup_update_leftovers()

    if not HAS_COM:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "缺少依赖",
            "未能加载 pywin32：\n%s\n\n请用带 pywin32 的 Python 运行：\n"
            "D:\\Python 3.12.4\\pythonw.exe" % COM_ERR)
        root.destroy()
        return

    inst = acquire_single_instance()
    if inst is None:
        try:
            notify_already_running()
        except Exception:
            pass
        return

    FloatWindow().run()


if __name__ == "__main__":
    main()
