# -*- coding: utf-8 -*-
"""
ui_layouts · 布局组件（导航、滚动容器、路径行、提示条等）
============================================================================
把「一组相关控件拼成一个完整区块」的样板代码封装出来。每个函数返回一个
可直接 pack/grid 的 widget。

公开 API：
    header_bar(parent, title, subtitle, F, status_var=, right_builder=, window=)
        顶部导航栏：品牌标记 + 标题 + 副标题 + 状态胶囊 + 右上角窗口按钮 + 拖动

    hint_bar(parent, text, F)
        底部提示条：浅灰分割线 + 一行说明

    section_label / page_label / muted_label
        各级文字标签的统一工厂

    ScrollableFrame(parent)
        可滚动容器（带滚轮支持）

    path_row(parent, label, var, on_change=, browse_title=)
        「标签 + 输入框 + 浏览…」一行，常用于路径输入
"""
import tkinter as tk
from tkinter import ttk, filedialog

from .ui_design import (
    C_PAGE, C_SURFACE, C_TEXT, C_MUTED, C_HINT, C_BORDER, C_ACCENT,
)
from .ui_window import (
    sp, enable_drag, make_borderless, WIN_MIN, WIN_MAX, WIN_CLOSE,
)
from .ui_render import logo_photo
from .ui_widgets import (
    primary_button, secondary_button, ghost_button, StatusPill,
    RoundedEntry, stat_card,
)


# ============================================================
# 顶部导航
# ============================================================
def header_bar(parent, title, subtitle='', F=None, status_var=None,
               right_builder=None, window=None):
    """页面顶部导航栏。

    结构：
        [logo]  title              [状态胶囊] [可选按钮] [—  ❐  ✕]
                subtitle
        ──────────────────────────────────────────
    按住 logo / title / subtitle 区域可拖动窗口（需传 window）。

    若传 window，会自动 make_borderless 并在右上角加最小化/最大化/关闭按钮。
    """
    try:
        parent.configure(bg=C_PAGE)
    except Exception:
        pass

    bar = tk.Frame(parent, bg=C_PAGE)
    row = tk.Frame(bar, bg=C_PAGE)
    row.pack(fill='x')

    # —— 左侧：logo + 标题 + 副标题 ——
    left = tk.Frame(row, bg=C_PAGE)
    left.pack(side='left', fill='x', expand=True, padx=(18, 8), pady=13)

    mark_sz = sp(34)
    mark = tk.Canvas(left, width=mark_sz, height=mark_sz, bg=C_PAGE,
                     highlightthickness=0, bd=0)
    mark.pack(side='left', padx=(0, 12))
    _mimg = logo_photo(34)
    mark._mimg = _mimg
    mark.create_image(0, 0, anchor='nw', image=_mimg)

    box = tk.Frame(left, bg=C_PAGE)
    box.pack(side='left', fill='x', expand=True)
    t_lbl = tk.Label(box, text=title, bg=C_PAGE, fg=C_TEXT,
                     font=(F.h1 if F else ('Segoe UI', 17, 'bold')))
    t_lbl.pack(anchor='w')
    s_lbl = None
    if subtitle:
        s_lbl = tk.Label(box, text=subtitle, bg=C_PAGE, fg=C_MUTED,
                         font=(F.small if F else ('Microsoft YaHei UI', 10)),
                         justify='left')
        s_lbl.pack(anchor='w', pady=(4, 0))

    # —— 右侧：状态胶囊 + 业务按钮 + 窗口按钮 ——
    right = tk.Frame(row, bg=C_PAGE)
    right.pack(side='right', padx=(8, 14), pady=13)

    if window is not None:
        make_borderless(window)
        ghost_button(right, '✕', command=lambda: WIN_CLOSE(window),
                     padx=11, pady=4).pack(side='right', padx=(2, 0))
        ghost_button(right, '❐', command=lambda: WIN_MAX(window),
                     padx=11, pady=4).pack(side='right', padx=(2, 0))
        ghost_button(right, '—', command=lambda: WIN_MIN(window),
                     padx=11, pady=4).pack(side='right', padx=(2, 0))
    if right_builder is not None:
        try:
            right_builder(right)
        except Exception:
            pass
    if status_var is not None:
        StatusPill(right, status_var, F).pack(side='right', padx=(8, 0))

    # —— 底部细分割线 ——
    tk.Frame(bar, bg=C_BORDER, height=1).pack(fill='x')

    # —— 绑定拖动 ——
    drag_targets = [bar, row, left, box, mark, t_lbl]
    if s_lbl is not None:
        drag_targets.append(s_lbl)
    if window is not None:
        enable_drag(window, *drag_targets)
    return bar


# ============================================================
# 底部提示 / 分割线
# ============================================================
def hint_bar(parent, text, F=None):
    """底部一行小灰字提示。"""
    f = tk.Frame(parent, bg=C_PAGE)
    tk.Frame(f, bg=C_BORDER, height=1).pack(fill='x')
    tk.Label(f, text=text, bg=C_PAGE, fg=C_HINT,
             font=(F.small if F else ('Microsoft YaHei UI', 10)),
             justify='left', wraplength=980).pack(anchor='w', padx=18,
                                                  pady=(10, 14))
    return f


# ============================================================
# 各级文字标签
# ============================================================
def section_label(parent, text, F=None):
    """分区小标题：粗体正文色。"""
    return tk.Label(parent, text=text, bg=C_PAGE, fg=C_TEXT,
                    font=(F.h2 if F else ('Microsoft YaHei UI', 12, 'bold')))


def page_label(parent, text, F=None, fg=None):
    """说明性文字：默认次要灰色。"""
    return tk.Label(parent, text=text, bg=C_PAGE, fg=(fg or C_MUTED),
                    font=(F.small if F else ('Microsoft YaHei UI', 10)))


def muted_label(parent, text, F=None, bg=None, wrap=None):
    """页面内的次要说明文字。"""
    return tk.Label(parent, text=text, bg=(bg or C_SURFACE), fg=C_MUTED,
                    font=(F.small if F else ('Microsoft YaHei UI', 10)),
                    justify='left', wraplength=(wrap or 0))


# ============================================================
# 可滚动容器
# ============================================================
class ScrollableFrame(tk.Frame):
    """带纵向滚动条的容器；滚轮在容器内自动接管。

    滚轮规则（来自原版）：
        - 鼠标当前在 ScrollableFrame 内部 → 滚动它
        - 鼠标在 Text/Treeview/Listbox → 不抢，让控件自己处理
        - Shift+滚轮 留给横向滚动（Treeview 自己处理）
    """
    def __init__(self, parent, bg=C_PAGE):
        tk.Frame.__init__(self, parent, bg=bg)
        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                                 yscrollincrement=sp(10))
        self._vbar = ttk.Scrollbar(self, orient='vertical',
                                   command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vbar.set)
        self._vbar.pack(side='right', fill='y')
        self._canvas.pack(side='left', fill='both', expand=True)
        self.container = tk.Frame(self._canvas, bg=bg)
        self._win = self._canvas.create_window((0, 0), window=self.container,
                                               anchor='nw')
        self.container.bind('<Configure>', lambda _e: self._refresh())
        self._canvas.bind('<Configure>', lambda _e: self._refresh())
        try:
            self._canvas.bind_all('<MouseWheel>', self._wheel, add='+')
            self._canvas.bind_all('<Button-4>', self._wheel, add='+')
            self._canvas.bind_all('<Button-5>', self._wheel, add='+')
        except Exception:
            pass
        for w in (self, self._canvas, self.container):
            w.bind('<Enter>', lambda _e: self._focus_canvas(), add='+')
        # 内容可能在打开后陆续完成布局 —— 多刷几次，保证一打开就能滚
        for d in (60, 200, 500, 1000):
            try:
                self.after(d, self._refresh)
            except Exception:
                pass

    def _focus_canvas(self):
        try:
            self._canvas.focus_set()
        except Exception:
            pass

    def _refresh(self):
        c = self._canvas
        try:
            cw = c.winfo_width()
            if cw <= 1:
                return
            ch = c.winfo_height()
            rh = self.container.winfo_reqheight()
            c.itemconfigure(self._win, width=cw)
            c.itemconfigure(self._win, height=(ch if rh < ch else rh))
            c.configure(scrollregion=(0, 0, cw, max(rh, ch)))
        except Exception:
            pass

    def _wheel(self, e):
        try:
            anc = self.winfo_containing(e.x_root, e.y_root)
        except Exception:
            return
        while anc is not None:
            if anc is self:
                break
            try:
                cls = anc.winfo_class()
            except Exception:
                return
            if cls in ('Text', 'Treeview', 'Listbox'):
                return
            anc = anc.master
        else:
            return
        if getattr(e, 'num', None) == 4:
            d = -1
        elif getattr(e, 'num', None) == 5:
            d = 1
        else:
            d = -1 if getattr(e, 'delta', 0) > 0 else 1
        try:
            self._canvas.yview_scroll(d * 3, 'units')
        except Exception:
            pass


# ============================================================
# 「标签 + 输入框 + 浏览…」一行
# ============================================================
def path_row(parent, label, var, on_change=None, browse_title='选择文件夹'):
    """构建「成绩文件夹：[____浏览…]」之类的输入行。

    on_change: 输入框失焦 / 回车时调用；浏览按钮确认后也会调。
    """
    fr = ttk.Frame(parent)
    fr.pack(fill='x', pady=5)
    ttk.Label(fr, text=label, width=10).pack(side='left')
    re_ = RoundedEntry(fr, textvariable=var)
    re_.pack(side='left', fill='x', expand=True, padx=(6, 8))
    ent = re_.entry
    if on_change is not None:
        ent.bind('<Return>', lambda _e: on_change())
        ent.bind('<FocusOut>', lambda _e: on_change())

    def browse():
        d = filedialog.askdirectory(title=browse_title)
        if d:
            var.set(d)
            if on_change is not None:
                on_change()

    secondary_button(fr, '浏览…', command=browse, padx=12, pady=6).pack(
        side='left')
    return fr


# ============================================================
# 任务页公共骨架（三页复用，减少样板代码）
# ============================================================
def page_scaffold(win, F, title, subtitle, stat_caption, status_var):
    """任务页公共骨架：顶部导航 + 滚动容器 + 顶部统计卡。

    返回 (host, var_big, var_side)：
        host     滚动容器内容区，后续各卡片 section 都 pack 到这里
        var_big  统计卡大数字的 StringVar
        var_side 统计卡右侧说明文字的 StringVar
    """
    var_big = tk.StringVar(value='0')
    var_side = tk.StringVar(value='')
    header_bar(win, title, subtitle, F, status_var=status_var,
               window=win).pack(fill='x')
    scroll = ScrollableFrame(win)
    scroll.pack(fill='both', expand=True)
    host = scroll.container
    stat_card(host, var_big, stat_caption, var_side, F).pack(
        fill='x', padx=12, pady=5)
    return host, var_big, var_side


def action_bar(parent, run_text, on_clear, preview_text=None,
               open_text='打开报告', progress='indeterminate'):
    """执行按钮行：可选「预览」+「开始」+ 打开结果 + 定位 + 清空日志 + 进度条。

    返回 (btn_prev, btn_run, btn_open, btn_rev, pb)：
        btn_prev  预览按钮；preview_text 为 None 时返回 None（任务①无预览）
        btn_run   开始按钮（主色）
        btn_open  打开结果按钮（初始 disabled，完成时由调用方 enable）
        btn_rev   在文件管理器定位按钮（初始 disabled）
        pb        进度条（位于 parent 内、按钮行下方）
    """
    bfr = ttk.Frame(parent)
    bfr.pack(fill='x')
    btn_prev = None
    if preview_text:
        btn_prev = secondary_button(bfr, preview_text, padx=14, pady=6)
        btn_prev.pack(side='left')
    btn_run = primary_button(bfr, run_text)
    btn_run.pack(side='left', padx=(6 if btn_prev else 0))
    btn_open = secondary_button(bfr, open_text, state='disabled',
                                padx=14, pady=6)
    btn_open.pack(side='left', padx=6)
    btn_rev = secondary_button(bfr, '在访达中显示', state='disabled',
                               padx=14, pady=6)
    btn_rev.pack(side='left', padx=6)
    ghost_button(bfr, '清空日志', command=on_clear,
                 padx=12, pady=6).pack(side='left', padx=6)
    pb = ttk.Progressbar(parent, mode=progress)
    pb.pack(fill='x', pady=(6, 0))
    return btn_prev, btn_run, btn_open, btn_rev, pb