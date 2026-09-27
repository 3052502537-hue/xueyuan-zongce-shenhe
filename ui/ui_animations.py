# -*- coding: utf-8 -*-
"""
ui_animations · 动效（克制型：仅在交互反馈点使用）
============================================================================
原则：
    - 全部用 after() 链式调度，不引入第三方库
    - 任何动效都不能阻塞主线程
    - 所有动效可被新事件打断（取消旧的 after）

公开 API：
    fade_in(win, ms=220, steps=12)
        窗口从透明到完全不透明（保留原版）

    animate_progress(pb, target, steps=6, delay=14)
        进度条平滑过渡到目标值（保留原版）

    stagger_highlight(cards, accent=, delay_between=, settle_ms=)
        依次让一组 RoundedCard 短暂亮起主色边框后回落（卡片入场动效）

    pulse(widget, color='#1890FF', count=2, duration=400)
        控件短暂高亮（引导注意）

    shake(widget, distance=4, count=3, duration=60)
        控件左右抖动（错误反馈）

    count_up(var, target, duration=400, steps=12)
        数字滚动动画（stat_card 数字变化时）
"""
import tkinter as tk

from .ui_design import C_ACCENT, C_BORDER


# ============================================================
# 通用工具
# ============================================================
class _Anim(object):
    """动效句柄：用于中途取消未完成的动效。"""

    def __init__(self):
        self._after_ids = []
        self._alive = True

    def after(self, widget, ms, fn):
        if not self._alive:
            return
        try:
            aid = widget.after(ms, fn)
            self._after_ids.append(aid)
        except Exception:
            pass

    def cancel(self, widget):
        self._alive = False
        for aid in self._after_ids:
            try:
                widget.after_cancel(aid)
            except Exception:
                pass
        self._after_ids.clear()


# ============================================================
# 窗口淡入（保留原版）
# ============================================================
def fade_in(win, ms=220, steps=12):
    """窗口透明度从 0 平滑过渡到 1。"""
    try:
        win.attributes('-alpha', 0.0)
    except Exception:
        return

    def step(i):
        try:
            # ease-out：越接近终点越慢，过渡更柔和
            t = i / float(steps)
            eased = 1.0 - (1.0 - t) ** 2
            win.attributes('-alpha', min(1.0, eased))
        except Exception:
            return
        if i < steps:
            try:
                win.after(max(1, int(ms / float(steps))),
                          lambda: step(i + 1))
            except Exception:
                pass

    step(1)


# ============================================================
# 卡片入场（依次亮起边框后回落）
# ============================================================
def stagger_highlight(cards, accent=None, delay_between=70, settle_ms=180,
                      settle=None):
    """依次让一组 RoundedCard 短暂亮起主色边框后回落，作为卡片入场动效。

    说明：Tk 的 Frame / Canvas 不支持 -alpha 透明度（只有 Tk/Toplevel 才有），
    因此不能像窗口那样做逐卡淡入；这里改为逐卡边框高亮，效果可见、不阻塞、
    也不影响布局。
    """
    accent = accent or C_ACCENT
    settle = settle or C_BORDER
    for i, card in enumerate(cards):

        def flash(c=card):
            try:
                c.recolor(border=accent)
                c.after(settle_ms, lambda cc=c: cc.recolor(border=settle))
            except Exception:
                pass

        try:
            card.after(i * delay_between, flash)
        except Exception:
            pass


# ============================================================
# 进度条平滑过渡（保留原版）
# ============================================================
def animate_progress(pb, target, steps=6, delay=14):
    """把 Progressbar 的值平滑过渡到 target（仅 determinate 模式有效）。"""
    try:
        cur = float(pb['value'])
    except Exception:
        cur = 0.0
    try:
        target = float(target)
    except Exception:
        return
    for i in range(1, steps + 1):
        # ease-out：前快后慢
        t = i / float(steps)
        eased = 1.0 - (1.0 - t) ** 2
        v = cur + (target - cur) * eased
        try:
            pb.after(int(delay * i), lambda vv=v: pb.configure(value=vv))
        except Exception:
            pass


# ============================================================
# 脉冲高亮（引导注意）
# ============================================================
def pulse(widget, color=None, count=2, duration=400):
    """让 widget 短暂闪烁一下（强调"该按这里"）。

    实现：临时把背景色在 原色↔主色 之间切换 count 次。
    """
    color = color or C_ACCENT
    try:
        original = widget.cget('bg')
    except Exception:
        original = None
    if original is None or not isinstance(original, str):
        return
    anim = _Anim()
    total = count * 2
    step_ms = max(40, duration // total)

    def step(i):
        if i > total:
            try:
                widget.configure(bg=original)
            except Exception:
                pass
            anim.cancel(widget)
            return
        try:
            widget.configure(bg=(color if i % 2 == 0 else original))
        except Exception:
            pass
        anim.after(widget, step_ms, lambda: step(i + 1))

    step(1)


# ============================================================
# 抖动（错误反馈）
# ============================================================
def shake(widget, distance=4, count=3, duration=60):
    """让 widget 左右抖动一下（错误反馈）。

    通过在 widget.master 上挪动 widget 实现的。widget 必须用 place 或
    已经 pack —— 这里只动 x，不动 y。
    """
    try:
        x0 = widget.winfo_x()
        y0 = widget.winfo_y()
    except Exception:
        return
    anim = _Anim()
    total = count * 2

    def step(i):
        if i > total:
            try:
                widget.place_configure(x=x0, y=y0)
            except Exception:
                pass
            anim.cancel(widget)
            return
        offset = distance if i % 2 == 0 else -distance
        try:
            widget.place_configure(x=x0 + offset, y=y0)
        except Exception:
            try:
                widget.place(x=x0 + offset, y=y0)
            except Exception:
                pass
        anim.after(widget, duration, lambda: step(i + 1))

    step(1)


# ============================================================
# 数字滚动（stat_card 数字变化时）
# ============================================================
def count_up(var, target, duration=400, steps=12, fmt=None,
             anchor_widget=None, on_done=None, from_value=None):
    """让 tk.StringVar 的值从某点平滑滚动到 target。

    var : tk.StringVar
    target : int / float（最终值）
    fmt : 格式化函数（默认 str）
    anchor_widget : 用于挂 after 的 widget（必填；通常是显示 var 的 Label）
    on_done : 动画完成后的回调（可选）
    from_value : 起点值；缺省则取 var.get() 的当前值

    ⚠ 当此函数被 var.trace_add('write') 回调调用时，var 已经被外部
        set 成 target 了，所以"起点"必须由调用方显式传进来。
    """
    if anchor_widget is None:
        try:
            var.set(str(target))
        except Exception:
            pass
        if on_done:
            try:
                on_done()
            except Exception:
                pass
        return
    try:
        target = float(target)
    except Exception:
        if on_done:
            try:
                on_done()
            except Exception:
                pass
        return

    if from_value is None:
        try:
            cur_raw = var.get().strip()
            cur = float(cur_raw) if cur_raw else 0.0
        except Exception:
            cur = 0.0
    else:
        try:
            cur = float(from_value)
        except Exception:
            cur = 0.0

    if fmt is None:
        fmt = lambda v: ('%d' % int(v)) if abs(v - int(v)) < 0.001 else ('%.1f' % v)

    delay = max(16, duration // steps)
    anim = _Anim()

    def step(i):
        t = i / float(steps)
        eased = 1.0 - (1.0 - t) ** 2
        v = cur + (target - cur) * eased
        try:
            var.set(fmt(v))
        except Exception:
            pass
        if i < steps:
            anim.after(anchor_widget, delay, lambda: step(i + 1))
        else:
            try:
                var.set(fmt(target))
            except Exception:
                pass
            if on_done:
                try:
                    on_done()
                except Exception:
                    pass

    step(0)