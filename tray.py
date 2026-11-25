# -*- coding: utf-8 -*-
"""极简系统托盘图标（纯 pywin32，不依赖 Pillow / pystray）。

为什么不用 pystray：它要拖 Pillow，打包体积从 12MB 涨到 31MB。
这里直接 Shell_NotifyIcon + 一个隐藏窗口收消息，功能一样，体积几乎为零。

用法：
    t = TrayIcon(ico_path, "提示文字", on_show=..., on_check=..., on_quit=...)
    t.start()   # 起线程跑消息循环
    t.stop()
回调都在**托盘线程**里执行，调用方要自己转回主线程。
"""
import os
import threading

import win32api
import win32con
import win32gui

WM_TRAY = win32con.WM_USER + 20
CMD_SHOW, CMD_CHECK, CMD_QUIT = 1001, 1002, 1003
CLASS_NAME = "CDRHelperTrayWnd"


class TrayIcon(object):
    def __init__(self, icon_path, tip="CorelDRAW 助手",
                 on_show=None, on_check=None, on_quit=None):
        self.icon_path = icon_path
        self.tip = tip
        self.on_show = on_show or (lambda: None)
        self.on_check = on_check or (lambda: None)
        self.on_quit = on_quit or (lambda: None)
        self.hwnd = 0
        self._hicon = 0
        self._thread = None
        self._ready = threading.Event()
        self._added = False

    # ---- 内部 -------------------------------------------------

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg == WM_TRAY:
            if lparam in (win32con.WM_LBUTTONUP, win32con.WM_LBUTTONDBLCLK):
                self._safe(self.on_show)
            elif lparam == win32con.WM_RBUTTONUP:
                self._menu()
            return 0
        if msg == win32con.WM_DESTROY:
            self._delete_icon(hwnd)
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    @staticmethod
    def _safe(fn):
        try:
            fn()
        except Exception:
            pass

    def _menu(self):
        try:
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu(menu, win32con.MF_STRING, CMD_SHOW, "显示助手")
            win32gui.AppendMenu(menu, win32con.MF_STRING, CMD_CHECK, "检查更新")
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, None)
            win32gui.AppendMenu(menu, win32con.MF_STRING, CMD_QUIT, "退出")
            pos = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.hwnd)
            cmd = win32gui.TrackPopupMenu(
                menu,
                win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON | win32con.TPM_RETURNCMD,
                pos[0], pos[1], 0, self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
            win32gui.DestroyMenu(menu)
            if cmd == CMD_SHOW:
                self._safe(self.on_show)
            elif cmd == CMD_CHECK:
                self._safe(self.on_check)
            elif cmd == CMD_QUIT:
                self._safe(self.on_quit)
        except Exception:
            pass

    def _add_icon(self, hwnd):
        size = win32api.GetSystemMetrics(win32con.SM_CXSMICON) or 16
        try:
            if self.icon_path and os.path.isfile(self.icon_path):
                self._hicon = win32gui.LoadImage(
                    0, self.icon_path, win32con.IMAGE_ICON, size, size,
                    win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
        except Exception:
            self._hicon = 0
        if not self._hicon:
            try:
                self._hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
            except Exception:
                self._hicon = 0
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        win32gui.Shell_NotifyIcon(
            win32gui.NIM_ADD,
            (hwnd, 0, flags, WM_TRAY, self._hicon, self.tip))
        self._added = True

    def _delete_icon(self, hwnd):
        if not self._added:
            return
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, 0))
        except Exception:
            pass
        self._added = False

    def _run(self):
        try:
            wc = win32gui.WNDCLASS()
            wc.lpszClassName = CLASS_NAME
            wc.hInstance = win32api.GetModuleHandle(None)
            wc.lpfnWndProc = self._wndproc
            win32gui.RegisterClass(wc)
        except Exception:
            pass                                   # 已注册过
        try:
            self.hwnd = win32gui.CreateWindow(
                CLASS_NAME, "CDRHelperTray", 0, 0, 0, 0, 0,
                0, 0, win32api.GetModuleHandle(None), None)
            self._add_icon(self.hwnd)
            self._ready.set()
            win32gui.PumpMessages()
        except Exception:
            self._ready.set()

    # ---- 对外 -------------------------------------------------

    def start(self):
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(3)

    def stop(self):
        try:
            if self.hwnd:
                win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
        self.hwnd = 0
