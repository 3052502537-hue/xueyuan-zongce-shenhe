# -*- coding: utf-8 -*-
"""
ui_window · 平台工具与窗口管理
============================================================================
跨平台兼容层 + 无边框窗口 + 拖动 + 置顶 + 屏幕居中放置 + 高 DPI 缩放。

公开 API：
    enable_dpi_awareness()        # Windows 下声明 DPI 感知（仅一次）
    sp(v)                          # 把设计像素换算成屏幕像素
    downloads_dir()                # 系统「下载」目录（Win 走注册表）
    default_out_root()             # 默认输出根目录（下载/班级成绩处理结果）
    open_path(p) / reveal_path(p)  # 用系统默认程序打开 / 在文件管理器中定位
    load_module / load_engine      # 按文件路径加载引擎（解决 exe 内嵌问题）
    make_borderless(win)           # 去掉系统标题栏（Win 改样式 / 其他 overrideredirect）
    enable_drag(win, *widgets)     # 按住指定区域拖动窗口
    place_window(win, w, h)        # 居中放置 + 适配屏幕 + minsize
    topmost_once(win, ms=800)      # 临时置顶拉前台
    force_foreground(win)          # 把窗口强制拉到前台
"""
import os
import sys
import ctypes
import platform
import subprocess
import importlib.util

import tkinter as tk

from .ui_design import IS_MAC, IS_WIN


# ============================================================
# 高 DPI
# ============================================================
def enable_dpi_awareness():
    """声明进程 DPI 感知，避免被 Windows 整体拉伸导致文字/图形发虚。

    Win 10+ 推荐 per-monitor（SetProcessDpiAwareness(2)），失败回退
    到 SetProcessDPIAware（Win 8.1+）；非 Win 直接 no-op。
    """
    if not IS_WIN or ctypes is None:
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# 进程内只需声明一次
if IS_WIN:
    enable_dpi_awareness()


# 全局缩放比例（96dpi = 1.0）
SCALE = 1.0
_SCALE_READY = False


def _init_scale(root):
    """按屏幕实际 DPI 计算缩放比例；必须在 root 创建后才能调用。"""
    global SCALE, _SCALE_READY
    try:
        SCALE = max(1.0, float(root.winfo_fpixels('1i')) / 96.0)
    except Exception:
        SCALE = 1.0
    _SCALE_READY = True
    return SCALE


def sp(v):
    """把设计尺寸换算成当前屏幕的像素尺寸。"""
    try:
        return int(round(float(v) * SCALE))
    except Exception:
        return v


def _sc_any(v):
    """递归把数值/容器里的所有数字按 SCALE 放大（用于安装到 tk.Misc 的 patch）。"""
    if isinstance(v, bool):
        return v
    if isinstance(v, (tuple, list)):
        return type(v)(_sc_any(x) for x in v)
    if isinstance(v, (int, float)):
        return int(round(v * SCALE))
    return v


_PATCHED = False


def _install_scaling_patches():
    """给 pack/grid 的 padx/pady、Treeview 列宽、Label wraplength 按比例放大。

    这样各页面里写死的 padx/pady/wraplength/列宽无需逐个改动，
    在高 DPI 下也能保持与设计一致的比例。
    """
    global _PATCHED
    if _PATCHED:
        return
    _PATCHED = True
    keys = ('padx', 'pady', 'ipadx', 'ipady')

    def wrap(fn):
        def w(self, cnf=None, **kw):
            for k in keys:
                if k in kw:
                    kw[k] = _sc_any(kw[k])
            return fn(self, cnf, **kw)
        return w

    for name in ('pack', 'pack_configure', 'grid', 'grid_configure'):
        try:
            setattr(tk.Misc, name, wrap(getattr(tk.Misc, name)))
        except Exception:
            pass

    try:
        _oc = tk.ttk.Treeview.column

        def _col(self, column, cnf=None, **kw):
            if 'width' in kw:
                kw['width'] = _sc_any(kw['width'])
            return _oc(self, column, cnf, **kw)
        tk.ttk.Treeview.column = _col
    except Exception:
        pass

    def wrap_init(cls):
        try:
            _oi = cls.__init__

            def _i(self, master=None, cnf=None, **kw):
                if 'wraplength' in kw:
                    kw['wraplength'] = _sc_any(kw['wraplength'])
                if isinstance(cnf, dict) and 'wraplength' in cnf:
                    cnf = dict(cnf)
                    cnf['wraplength'] = _sc_any(cnf['wraplength'])
                if cnf is None:
                    return _oi(self, master, **kw)
                return _oi(self, master, cnf, **kw)
            cls.__init__ = _i
        except Exception:
            pass

    wrap_init(tk.Label)
    wrap_init(tk.ttk.Label)


# ============================================================
# 系统路径
# ============================================================
def downloads_dir():
    """返回系统「下载」目录。

    Windows 下读注册表（用户可能改过位置）；macOS/Linux 直接探测家目录下的
    Downloads / 下载。失败兜底为 home。
    """
    home = os.path.expanduser('~')
    if IS_WIN:
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r'Software\Microsoft\Windows\CurrentVersion\Explorer'
                r'\User Shell Folders')
            val, _ = winreg.QueryValueEx(
                key, '{374DE290-123F-4565-9164-39C4925E467B}')
            winreg.CloseKey(key)
            val = os.path.expandvars(val)
            if val and os.path.isdir(val):
                return val
        except Exception:
            pass
        d = os.path.join(home, 'Downloads')
        return d if os.path.isdir(d) else home
    for c in (os.path.join(home, 'Downloads'), os.path.join(home, '下载')):
        if os.path.isdir(c):
            return c
    return home


def default_out_root():
    """默认产物输出根目录：~/Downloads/班级成绩处理结果。"""
    return os.path.join(downloads_dir(), '班级成绩处理结果')


def open_path(p):
    """用系统默认程序打开文件/文件夹。"""
    try:
        if IS_MAC:
            subprocess.run(['open', p], check=False)
        elif IS_WIN:
            os.startfile(p)   # noqa: S606
        else:
            subprocess.run(['xdg-open', p], check=False)
        return True
    except Exception:
        return False


def reveal_path(p):
    """在文件管理器中定位（macOS 选中 / Win 资源管理器高亮 / 其他只打开）。"""
    try:
        if IS_MAC:
            subprocess.run(['open', '-R', p], check=False)
        elif IS_WIN:
            subprocess.run(['explorer', '/select,',
                            os.path.normpath(p)], check=False)
        return True
    except Exception:
        return False


# ============================================================
# 引擎加载（exe 内嵌场景）
# ============================================================
def load_module(path, name):
    """按文件路径加载一个 Python 模块并注册到 sys.modules。"""
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def load_engine(here, filename, name):
    """在 here/ 或 its parent/ 找引擎文件并加载。

    PyInstaller 单文件 exe 把引擎作为 datas 内嵌到 _MEIPASS；
    源码运行时 here == parent，所以两个候选路径都要试。
    """
    cands = [os.path.join(here, filename),
             os.path.join(os.path.dirname(here), filename)]
    for p in cands:
        if os.path.exists(p):
            return load_module(p, name), p
    raise FileNotFoundError('未找到引擎「%s」。\n已查找：\n  %s'
                            % (filename, '\n  '.join(cands)))


# ============================================================
# 无边框窗口 / 拖动
# ============================================================
def _hwnd(win):
    win.update_idletasks()
    return ctypes.windll.user32.GetParent(win.winfo_id())


def make_borderless(win):
    """去掉系统标题栏（用自定义头部替代）。Win 改样式，其他平台 overrideredirect。"""
    if not IS_WIN or ctypes is None:
        try:
            win.overrideredirect(True)
        except Exception:
            pass
        return
    try:
        GWL_STYLE = -16
        WS_CAPTION = 0x00C00000
        u = ctypes.windll.user32
        hwnd = _hwnd(win)
        style = u.GetWindowLongW(hwnd, GWL_STYLE)
        new = style & ~WS_CAPTION
        if new != style:
            u.SetWindowLongW(hwnd, GWL_STYLE, new)
        u.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0027)
    except Exception:
        pass


def enable_drag(win, *widgets):
    """按住指定区域可拖动窗口（Win 走 ReleaseCapture API，其他走事件）。"""
    if IS_WIN and ctypes is not None:
        def _start(_e=None):
            try:
                u = ctypes.windll.user32
                u.ReleaseCapture()
                u.SendMessageW(_hwnd(win), 0x00A1, 0x2, 0)
            except Exception:
                pass
        for w in widgets:
            try:
                w.bind('<Button-1>', _start, add='+')
            except Exception:
                pass
        return
    st = {'x': 0, 'y': 0}

    def _press(e):
        st['x'] = e.x_root - win.winfo_x()
        st['y'] = e.y_root - win.winfo_y()

    def _move(e):
        try:
            win.geometry('+%d+%d' % (e.x_root - st['x'], e.y_root - st['y']))
        except Exception:
            pass

    for w in widgets:
        try:
            w.bind('<Button-1>', _press, add='+')
            w.bind('<B1-Motion>', _move, add='+')
        except Exception:
            pass


def _win_min(win):
    try:
        win.iconify()
    except Exception:
        pass


def _win_max(win):
    try:
        win.state('normal' if win.state() == 'zoomed' else 'zoomed')
    except Exception:
        pass


def _win_close(win):
    try:
        win.destroy()
    except Exception:
        pass


# 右上角最小化/最大化/关闭按钮（被 header_bar 调用）
WIN_MIN   = _win_min
WIN_MAX   = _win_max
WIN_CLOSE = _win_close


# ============================================================
# 放置 / 置顶
# ============================================================
def place_window(win, pref_w, pref_h, min_w=760, min_h=520):
    """按屏幕尺寸居中放置窗口，并设 minsize。

    规则：
        - 宽度  ∈ [min_w, screen*0.94]，初始为 pref_w
        - 高度  ∈ [min_h, screen*0.96]，初始为 pref_h
        - 最大不超过 screen×0.98
        - 垂直方向略偏上 -24px（视觉上更舒服）
    """
    try:
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
    except Exception:
        sw, sh = 1366, 768
    pw, ph = sp(pref_w), sp(pref_h)
    mw, mh = sp(min_w), sp(min_h)
    w = max(mw, min(pw, int(sw * 0.94)))
    h = max(mh, min(ph, int(sh * 0.90)))
    w = min(w, int(sw * 0.98))
    h = min(h, int(sh * 0.96))
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2 - sp(24))
    try:
        win.geometry('%dx%d+%d+%d' % (w, h, x, y))
        win.minsize(min(mw, w), min(mh, h))
    except Exception:
        pass


def force_foreground(win):
    """把窗口强制拉到前台并取得焦点（Windows 上滚轮事件依赖焦点）。"""
    try:
        win.lift()
    except Exception:
        pass
    if IS_WIN and ctypes is not None:
        try:
            u = ctypes.windll.user32
            hwnd = _hwnd(win)
            u.BringWindowToTop(hwnd)
            u.SetForegroundWindow(hwnd)
            u.SetFocus(hwnd)
        except Exception:
            pass
    try:
        win.focus_force()
    except Exception:
        pass


def topmost_once(win, ms=800):
    """临时置顶 ms 毫秒拉前台，然后取消置顶（仍保留前台状态）。"""
    try:
        win.attributes('-topmost', True)
        win.after(ms, lambda: win.attributes('-topmost', False))
    except Exception:
        pass
    force_foreground(win)
    try:
        win.after(150, lambda: force_foreground(win))
    except Exception:
        pass