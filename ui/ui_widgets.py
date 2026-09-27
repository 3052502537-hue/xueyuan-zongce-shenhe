# -*- coding: utf-8 -*-
"""
ui_widgets · 自定义控件
============================================================================
对 tkinter 控件做现代化包装：圆角按钮、圆角卡片、圆角勾选框、圆角输入框、
状态胶囊、表格（含行 hover、横向滚动）。

设计原则：
    - 不破坏 tk 控件的 API：configure / config / cget 全部可用
    - 所有控件用 Canvas + Pillow 预渲染图片实现圆角与阴影
    - 所有控件在 disabled 状态自动变灰

公开 API：
    primary_button / secondary_button / ghost_button
    RoundedCard / card_section / stat_card
    CheckBox / checkbox
    RoundedEntry / entry
    StatusPill
    scroll_table / enable_tree_hover
"""
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from .ui_design import (
    C_ACCENT, C_ACCENT_HOVER, C_ACCENT_SOFT, C_ACCENT_TINT,
    C_SURFACE, C_BORDER, C_BORDER_STRONG, C_BORDER_HOVER,
    C_TEXT, C_MUTED, C_HINT, C_SUBTLE, C_CARD_HOVER,
    C_TREE_HEAD, C_DISABLED_BG, C_DISABLED_FG,
    C_OK, C_WARN, C_ERR, RADIUS_BTN, RADIUS_CARD, RADIUS_INPUT,
    BG_OK_SOFT, BG_ERR_SOFT,
    Fonts,
)
from .ui_window import sp
from .ui_render import btn_photo, pill_photo, card_photo


# ============================================================
# 字体工具
# ============================================================
_BTN_FONT = None


def _btn_font():
    global _BTN_FONT
    if _BTN_FONT is None:
        _BTN_FONT = tkfont.Font(family=Fonts().fam, size=10)
    return _BTN_FONT


# ============================================================
# 圆角按钮
# ============================================================
class AnimButton(tk.Canvas):
    """圆角按钮（Canvas + Pillow 预渲染）。

    支持 4 种配色：
        primary   主色蓝
        secondary 次要灰
        ghost     透明（无边框）
        danger    错误红

    动效（克制型）：
        - press 时文字下沉 1px + 填充变深
        - hover 时切换 hover 底色与边框
    """
    _PALETTE = {
        'primary':   dict(bg=C_ACCENT,  hover=C_ACCENT_HOVER, pressed='#0855C4',
                          fg='#FFFFFF', border=C_ACCENT,
                          hover_border=C_ACCENT_HOVER),
        'secondary': dict(bg=C_SURFACE, hover=C_ACCENT_TINT,   pressed='#E1F0FF',
                          fg=C_TEXT, border=C_BORDER_STRONG, hover_border=C_ACCENT),
        'ghost':     dict(bg=C_SURFACE, hover=C_SUBTLE,        pressed='#EDF0F4',
                          fg=C_MUTED, border=None, hover_border=None),
        'danger':    dict(bg=C_ERR,     hover='#D9363E',       pressed='#C22F35',
                          fg='#FFFFFF', border=C_ERR, hover_border='#D9363E'),
    }

    def __init__(self, parent, text, command=None, kind='primary', **kw):
        p = dict(self._PALETTE.get(kind, self._PALETTE['secondary']))
        self._p = p
        self._text = text
        self._cmd = command
        self._disabled = False
        self._hovered = False
        self._pressed = False          # v3：按下状态
        self._padx = sp(kw.pop('padx', 16))
        self._pady = sp(kw.pop('pady', 8))
        self._radius = sp(kw.pop('radius', RADIUS_BTN))
        state = kw.pop('state', 'normal')
        surface = kw.pop('bg', C_SURFACE)
        f = _btn_font()
        w = f.measure(text) + 2 * self._padx
        h = f.metrics('linespace') + 2 * self._pady
        tk.Canvas.__init__(self, parent, width=w, height=h, bg=surface,
                           highlightthickness=0, bd=0, cursor='hand2')
        self._img_id = None
        self._txt_id = None
        self._photo = None
        self._render()
        self.bind('<Enter>', self._on_enter, add='+')
        self.bind('<Leave>', self._on_leave, add='+')
        self.bind('<ButtonPress-1>', self._on_press, add='+')
        self.bind('<ButtonRelease-1>', self._on_rel, add='+')
        if str(state) == 'disabled':
            self._disabled = True
            self._render()

    def _colors(self):
        """返回当前 (fill, border, fg)。"""
        if self._disabled:
            return C_DISABLED_BG, C_DISABLED_BG, C_DISABLED_FG
        border = self._p.get('hover_border', self._p['border'])
        if self._pressed:
            return (self._p.get('pressed', self._p['hover']),
                    (border or self._p['hover']), self._p['fg'])
        if self._hovered:
            return self._p['hover'], (border or self._p['hover']), self._p['fg']
        return self._p['bg'], (self._p['border'] or self._p['bg']), self._p['fg']

    def _render(self):
        w = int(self['width'])
        h = int(self['height'])
        fill, border, fg = self._colors()
        photo = btn_photo(w, h, self._radius, fill, border)
        self._photo = photo
        off = 1 if self._pressed else 0   # 按下时文字下沉 1px
        cx, cy = w / 2.0 + off, h / 2.0 + off
        if self._img_id is None:
            self._img_id = self.create_image(0, 0, anchor='nw', image=photo)
            self._txt_id = self.create_text(cx, cy, text=self._text,
                                            fill=fg, font=_btn_font())
        else:
            self.itemconfigure(self._img_id, image=photo)
            self.itemconfigure(self._txt_id, text=self._text, fill=fg)
            self.coords(self._txt_id, cx, cy)

    # —— 交互回调 ——
    def _on_enter(self, _e):
        if not self._disabled:
            self._hovered = True
            self._render()

    def _on_leave(self, _e):
        if not self._disabled:
            self._hovered = False
            self._render()

    def _on_press(self, _e):
        if not self._disabled:
            self._pressed = True
            self._hovered = True
            self._render()

    def _on_rel(self, _e):
        if self._disabled:
            return
        self._pressed = False
        # 鼠标可能仍在按钮上（_on_enter 已经设了 _hovered=True），保持；
        # 若鼠标已离开（_on_leave 设了 _hovered=False），保持 False。
        self._render()
        if self._cmd is not None:
            try:
                self._cmd()
            except Exception:
                pass

    def configure(self, cnf=None, **kw):
        if cnf is not None and isinstance(cnf, dict):
            kw = dict(cnf, **kw)
        dirty = False
        if 'text' in kw:
            self._text = kw.pop('text')
            dirty = True
        if 'command' in kw:
            self._cmd = kw.pop('command')
        if 'state' in kw:
            self._disabled = (str(kw.pop('state')) == 'disabled')
            try:
                tk.Canvas.configure(
                    self, cursor=('arrow' if self._disabled else 'hand2'))
            except Exception:
                pass
            dirty = True
        if kw:
            try:
                tk.Canvas.configure(self, **kw)
            except Exception:
                pass
        if dirty:
            f = _btn_font()
            w = f.measure(self._text) + 2 * self._padx
            h = f.metrics('linespace') + 2 * self._pady
            if int(self['width']) != w or int(self['height']) != h:
                try:
                    tk.Canvas.configure(self, width=w, height=h)
                except Exception:
                    pass
            self._render()

    config = configure


def primary_button(parent, text, command=None):
    return AnimButton(parent, text, command=command, kind='primary')


def secondary_button(parent, text, command=None, **kw):
    return AnimButton(parent, text, command=command, kind='secondary', **kw)


def ghost_button(parent, text, command=None, **kw):
    return AnimButton(parent, text, command=command, kind='ghost', **kw)


# ============================================================
# 状态胶囊（圆角文字标签，自动按文字变色）
# ============================================================
class StatusPill(tk.Canvas):
    """根据文字自动选择配色：
        含「完成 / ✔ / 一致」 → 绿色
        含「出错 / 失败 / ✘ / 未通过」 → 红色
        其它 → 主色蓝
    """
    def __init__(self, parent, var, F=None):
        self._var = var
        self._font = (F.small if F else _btn_font())
        self._bg = C_ACCENT_SOFT
        self._fg = C_ACCENT
        tk.Canvas.__init__(self, parent, bg=C_SURFACE, highlightthickness=0,
                           bd=0)
        self._img_id = None
        self._txt_id = None
        self._photo = None
        try:
            var.trace_add('write', lambda *a: self._update())
        except Exception:
            pass
        self._update()

    def _update(self):
        t = str(self._var.get())
        if any(k in t for k in ('完成', '✔', '一致')):
            self._bg, self._fg = BG_OK_SOFT, C_OK
        elif any(k in t for k in ('出错', '失败', '✘', '未通过')):
            self._bg, self._fg = BG_ERR_SOFT, C_ERR
        else:
            self._bg, self._fg = C_ACCENT_SOFT, C_ACCENT
        self._render()

    def _render(self):
        txt = str(self._var.get())
        f = self._font
        w = f.measure(txt) + sp(26)
        h = f.metrics('linespace') + sp(12)
        try:
            self.configure(width=w, height=h)
        except Exception:
            return
        photo = pill_photo(w, h, self._bg)
        self._photo = photo
        if self._img_id is None:
            self._img_id = self.create_image(0, 0, anchor='nw', image=photo)
            self._txt_id = self.create_text(w / 2.0, h / 2.0, text=txt,
                                            fill=self._fg, font=f)
        else:
            self.itemconfigure(self._img_id, image=photo)
            self.itemconfigure(self._txt_id, text=txt, fill=self._fg)


# ============================================================
# 圆角卡片（最常用的容器）
# ============================================================
class RoundedCard(tk.Canvas):
    """圆角带阴影的容器，内部嵌入一个 tk.Frame 放实际内容。

    recolor(fill=, border=) 可在 hover 时切换底色（如首页卡片）。

    v3 微调：
        - recolor 改色时 inner frame 也同步着色（避免内容背景与卡片不一致）
        - 多触发几次 _pump() 应对 DPI 缩放后尺寸回填
    """
    def __init__(self, parent, radius=RADIUS_CARD, pad=18, fill=C_SURFACE,
                 border=C_BORDER, bg=C_SURFACE, shadow=True):
        tk.Canvas.__init__(self, parent, bg=bg, highlightthickness=0, bd=0,
                           height=10)
        self._radius = sp(radius)
        self._pad = sp(pad)
        self._fill = fill
        self._border = border
        self._shadow = shadow
        self._m = sp(7) if shadow else 0
        self._img_id = None
        self._photo = None
        self.inner = tk.Frame(self, bg=fill)
        off = self._off()
        self._win = self.create_window((off, off), window=self.inner,
                                       anchor='nw')
        self.bind('<Configure>', self._on_canvas)
        self.inner.bind('<Configure>', self._on_inner)
        # 多触发几次 _pump()，应对某些系统首次 layout 失败
        for d in (20, 80, 200, 500):
            try:
                self.after(d, self._pump)
            except Exception:
                pass

    def _off(self):
        return self._m + self._pad

    def recolor(self, fill=None, border=None):
        """切换卡片底色与边框色；hover 用。"""
        if fill:
            self._fill = fill
            try:
                self.inner.configure(bg=fill)
            except Exception:
                pass
        if border:
            self._border = border
        self._redraw()

    def _sync_height(self):
        try:
            h = self.inner.winfo_reqheight() + 2 * self._off()
            if int(self['height']) != h:
                tk.Canvas.configure(self, height=h)
        except Exception:
            pass

    def _pump(self):
        try:
            w = self.winfo_width()
            if w <= 3:
                return
            ih = max(1, self.inner.winfo_reqheight())
            self.itemconfigure(self._win, width=max(1, w - 2 * self._off()),
                               height=ih)
        except Exception:
            pass
        self._sync_height()
        self._redraw()

    def _on_inner(self, _e):
        self._sync_height()

    def _on_canvas(self, _e):
        self._pump()

    def _redraw(self):
        try:
            w = int(self.winfo_width())
            h = int(self.winfo_height())
        except Exception:
            return
        if w <= 3 or h <= 3:
            return
        photo = card_photo(w, h, self._radius, self._fill, self._border,
                           self._shadow)
        self._photo = photo
        if self._img_id is None:
            self._img_id = self.create_image(0, 0, anchor='nw', image=photo)
        else:
            self.itemconfigure(self._img_id, image=photo)
        try:
            self.tag_lower(self._img_id)
        except Exception:
            pass


class _CardBody(tk.Frame):
    """card_section 返回的 body：pack() 直接转发到外层 Card。"""
    def __init__(self, card, **kw):
        tk.Frame.__init__(self, card.inner, **kw)
        self._card = card

    def pack(self, **kw):
        self._card.pack(**kw)
        return self

    def pack_configure(self, **kw):
        self._card.pack_configure(**kw)
        return self


class _CardPacker(object):
    """stat_card 返回的轻量代理：转 pack 调用给 Card。"""
    def __init__(self, card):
        self._card = card

    def pack(self, **kw):
        self._card.pack(**kw)
        return self

    def pack_configure(self, **kw):
        self._card.pack_configure(**kw)
        return self


def card_section(parent, title, hint=None, F=None, radius=RADIUS_CARD,
                 pad=18, shadow=True):
    """带标题的圆角卡片分区：左侧蓝色竖条 + 标题 + 可选提示。"""
    card = RoundedCard(parent, radius=radius, pad=pad, shadow=shadow)
    hdr = tk.Frame(card.inner, bg=C_SURFACE)
    tk.Frame.pack(hdr, fill='x', pady=(0, 10))
    tk.Frame(hdr, bg=C_ACCENT, width=sp(3), height=sp(15)).pack(
        side='left', pady=1)
    tk.Label(hdr, text=title, bg=C_SURFACE, fg=C_TEXT,
             font=(F.sec if F else ('Microsoft YaHei UI', 11, 'bold'))).pack(
        side='left', padx=(8, 0))
    if hint:
        tk.Label(hdr, text=hint, bg=C_SURFACE, fg=C_MUTED,
                 font=(F.small if F else ('Microsoft YaHei UI', 10))).pack(
            side='left', padx=(10, 0))
    body = _CardBody(card, bg=C_SURFACE)
    tk.Frame.pack(body, fill='both', expand=True)
    return body


def stat_card(parent, num_var, caption, side_var=None, F=None,
              radius=RADIUS_CARD):
    """统计卡片：大数字 + 副标题（+ 可选右侧说明文字）。

    v3 动效：当 num_var 的值从合法数字变为另一个合法数字时，
    自动触发数字滚动动画（调用方无需感知）；首次设置值与非法值不进动画。
    """
    card = RoundedCard(parent, radius=radius, pad=20)
    row = tk.Frame(card.inner, bg=C_SURFACE)
    tk.Frame.pack(row, fill='x')
    left = tk.Frame(row, bg=C_SURFACE)
    left.pack(side='left', anchor='n')
    num_lbl = tk.Label(left, textvariable=num_var, bg=C_SURFACE, fg=C_ACCENT,
                      font=(F.stat if F else ('Segoe UI', 34, 'bold')))
    num_lbl.pack(anchor='w')
    tk.Label(left, text=caption, bg=C_SURFACE, fg=C_MUTED,
             font=(F.small if F else ('Microsoft YaHei UI', 10))).pack(
        anchor='w', pady=(2, 0))
    if side_var is not None:
        right = tk.Frame(row, bg=C_SURFACE)
        right.pack(side='right', anchor='e')
        tk.Label(right, textvariable=side_var, bg=C_SURFACE, fg=C_MUTED,
                 font=(F.small if F else ('Microsoft YaHei UI', 10)),
                 justify='right', anchor='e').pack(anchor='e')

    # —— 数字滚动动画（仅当目标值是合法数字时才动效）——
    # ⚠ 注意：
    #   1) count_up 内部的 var.set() 也会触发这个 trace，所以需要
    #      _animating 标志位防止递归启动新动画。
    #   2) trace 触发时 var 已经被外部 set 成 target，所以 count_up 必须
    #      从 from_value=displayed 起步；否则会从 100→100（瞬时跳）。
    #   3) StringVar(value='0') 构造时不触发 trace，因此 displayed 必须
    #      从 num_var.get() 当前值初始化，否则首次 set 会被当作「首次记下」
    #      而跳过动画。
    try:
        from .ui_animations import count_up
        try:
            init_raw = num_var.get().strip()
            init_val = float(init_raw) if init_raw else None
        except Exception:
            init_val = None
        displayed = {'v': init_val}
        animating = {'on': False}

        def _to_float(s):
            try:
                return float(str(s).strip())
            except Exception:
                return None

        def _on_num_change(*_a):
            if animating['on']:
                return
            new = _to_float(num_var.get())
            if new is None:
                displayed['v'] = None
                return
            if displayed['v'] is None:
                # 首次见到数字：直接记下，不动效
                displayed['v'] = new
                return
            if new != displayed['v']:
                # 真正的目标变化：从 displayed['v'] 平滑到 new
                animating['on'] = True
                try:
                    count_up(num_var, new, duration=380, steps=12,
                             anchor_widget=num_lbl, from_value=displayed['v'],
                             on_done=lambda: animating.__setitem__('on', False))
                except Exception:
                    animating['on'] = False
                displayed['v'] = new

        num_var.trace_add('write', _on_num_change)
    except Exception:
        pass

    return _CardPacker(card)


# ============================================================
# 表格（带横/纵滚动条 + 行 hover）
# ============================================================
def enable_tree_hover(tree):
    """为 Treeview 启用行 hover 高亮（背景色变 C_SUBTLE）。"""
    state = {'row': ''}
    try:
        tree.tag_configure('hover', background=C_SUBTLE)
    except Exception:
        return

    def _clear(row):
        if row:
            try:
                tree.item(row, tags=())
            except Exception:
                pass

    def _on_motion(e):
        row = tree.identify_row(e.y)
        if row == state['row']:
            return
        _clear(state['row'])
        if row:
            try:
                tree.item(row, tags=('hover',))
            except Exception:
                pass
        state['row'] = row

    def _on_leave(_e):
        _clear(state['row'])
        state['row'] = ''

    tree.bind('<Motion>', _on_motion, add='+')
    tree.bind('<Leave>', _on_leave, add='+')


def scroll_table(parent, columns, height=10):
    """创建带「纵向 + 横向」滚动条的表格容器，返回 (容器, 表格)。

    高 DPI 下列较多时列会被截断，横向滚动条可让鼠标快速定位到右侧列。
    Shift+滚轮 = 横向滚动。
    """
    wrap = ttk.Frame(parent)
    tree = ttk.Treeview(wrap, columns=columns, show='headings', height=height)
    vs = ttk.Scrollbar(wrap, orient='vertical', command=tree.yview)
    hs = ttk.Scrollbar(wrap, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vs.grid(row=0, column=1, sticky='ns')
    hs.grid(row=1, column=0, sticky='ew')
    wrap.rowconfigure(0, weight=1)
    wrap.columnconfigure(0, weight=1)
    enable_tree_hover(tree)
    tree.bind('<Shift-MouseWheel>',
              lambda e: tree.xview_scroll(-1 if e.delta > 0 else 1, 'units'),
              add='+')
    return wrap, tree


# ============================================================
# 圆角勾选框
# ============================================================
class CheckBox(tk.Frame):
    """圆角勾选框：勾选时显示主色蓝方块 + 白色对勾。"""
    def __init__(self, parent, text='', variable=None, command=None, F=None,
                 bg=C_SURFACE, size=18):
        tk.Frame.__init__(self, parent, bg=bg, cursor='hand2')
        self._var = (variable if variable is not None
                     else tk.BooleanVar(value=False))
        self._cmd = command
        self._bg = bg
        self._size = sp(size)
        self._disabled = False
        self._hover = False
        self._font = F.small if F else _btn_font()
        self._cv = tk.Canvas(self, width=self._size, height=self._size, bg=bg,
                             highlightthickness=0, bd=0, cursor='hand2')
        self._cv.pack(side='left')
        self._lbl = tk.Label(self, text=text, bg=bg, fg=C_TEXT,
                             font=self._font, cursor='hand2')
        self._lbl.pack(side='left', padx=(8, 0))
        self._img_id = None
        self._txt_id = None
        self._photo = None
        self._render()
        for w in (self, self._cv, self._lbl):
            w.bind('<Button-1>', self._toggle, add='+')
            w.bind('<Enter>', self._enter, add='+')
            w.bind('<Leave>', self._leave, add='+')
        try:
            self._var.trace_add('write', lambda *a: self._render())
        except Exception:
            pass

    def _render(self):
        on = bool(self._var.get())
        if self._disabled:
            box_bg, border = C_DISABLED_BG, C_BORDER_STRONG
        elif on:
            box_bg, border = C_ACCENT, C_ACCENT
        else:
            box_bg = C_SURFACE
            border = C_ACCENT if self._hover else C_BORDER_STRONG
        s = self._size
        photo = btn_photo(s, s, sp(5), box_bg, border)
        self._photo = photo
        tick = '✓' if on else ''
        fg = '#FFFFFF' if on else C_TEXT
        if self._img_id is None:
            self._img_id = self._cv.create_image(0, 0, anchor='nw', image=photo)
            self._txt_id = self._cv.create_text(s / 2.0, s / 2.0, text=tick,
                                                fill=fg, font=_btn_font())
        else:
            self._cv.itemconfigure(self._img_id, image=photo)
            self._cv.itemconfigure(self._txt_id, text=tick, fill=fg)
        try:
            self._lbl.configure(
                fg=(C_DISABLED_FG if self._disabled else C_TEXT))
        except Exception:
            pass

    def _enter(self, _e):
        if not self._disabled:
            self._hover = True
            self._render()

    def _leave(self, _e):
        if not self._disabled:
            self._hover = False
            self._render()

    def _toggle(self, _e=None):
        if self._disabled:
            return
        try:
            self._var.set(not bool(self._var.get()))
        except Exception:
            return
        if self._cmd is not None:
            try:
                self._cmd()
            except Exception:
                pass

    def configure(self, cnf=None, **kw):
        if cnf is not None and isinstance(cnf, dict):
            kw = dict(cnf, **kw)
        if 'state' in kw:
            self._disabled = (str(kw.pop('state')) == 'disabled')
            try:
                cur = 'arrow' if self._disabled else 'hand2'
                for w in (self, self._cv, self._lbl):
                    w.configure(cursor=cur)
            except Exception:
                pass
            self._render()
        if 'text' in kw:
            try:
                self._lbl.configure(text=kw.pop('text'))
            except Exception:
                pass
        if kw:
            try:
                tk.Frame.configure(self, **kw)
            except Exception:
                pass

    config = configure


def checkbox(parent, text='', variable=None, command=None, F=None, **kw):
    return CheckBox(parent, text=text, variable=variable, command=command,
                    F=F, **kw)


# ============================================================
# 圆角输入框
# ============================================================
class RoundedEntry(tk.Frame):
    """圆角输入框：聚焦时边框变主色蓝。"""
    def __init__(self, parent, textvariable=None, width=None, font=None,
                 height=32, radius=RADIUS_INPUT, bg=C_SURFACE):
        tk.Frame.__init__(self, parent, bg=bg)
        self._bg = bg
        self._height = sp(height)
        self._radius = sp(radius)
        self._focused = False
        self._img_id = None
        self._photo = None
        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             height=self._height)
        self._cv.pack(fill='both', expand=True)
        self.entry = tk.Entry(self, textvariable=textvariable, relief='flat',
                              bd=0, highlightthickness=0, bg=C_SURFACE,
                              fg=C_TEXT, insertbackground=C_ACCENT,
                              font=(font or _btn_font()),
                              disabledbackground=C_SURFACE)
        self._win = self._cv.create_window(sp(12), self._height / 2.0,
                                           anchor='w', window=self.entry)
        self._cv.bind('<Configure>', self._on_cfg)
        self.entry.bind('<FocusIn>', self._on_focus, add='+')
        self.entry.bind('<FocusOut>', self._on_blur, add='+')
        self._cv.bind('<Button-1>', lambda _e: self.entry.focus_set(), add='+')
        if width:
            try:
                self.entry.configure(width=width)
            except Exception:
                pass

    def _on_focus(self, _e):
        self._focused = True
        self._redraw()

    def _on_blur(self, _e):
        self._focused = False
        self._redraw()

    def _on_cfg(self, e):
        border = C_ACCENT if self._focused else C_BORDER_STRONG
        photo = btn_photo(max(2, e.width), max(2, e.height), self._radius,
                          C_SURFACE, border)
        self._photo = photo
        if self._img_id is None:
            self._img_id = self._cv.create_image(0, 0, anchor='nw', image=photo)
            self._cv.tag_lower(self._img_id)
        else:
            self._cv.itemconfigure(self._img_id, image=photo)
        try:
            self._cv.itemconfigure(self._win, width=max(1, e.width - sp(24)),
                                   height=max(1, e.height - sp(12)))
        except Exception:
            pass

    def _redraw(self):
        w = self._cv.winfo_width()
        h = self._cv.winfo_height()
        if w > 2:
            self._on_cfg(type('E', (), {'width': w, 'height': h})())


def entry(parent, textvariable=None, **kw):
    return RoundedEntry(parent, textvariable=textvariable, **kw)