# -*- coding: utf-8 -*-
"""
综测核算程序 · 统一入口
============================================================================
打开后先选任务，再进入对应页面：

  ① 综测两表核对    —— 详情表 vs 测评表 一致性核验（输出核验报告 xlsx）
  ② 挂科/附加分筛查  —— 必修课挂科 + 100分附加分（输出《挂科与附加分表》xlsx）
  ③ 回填成绩表      —— 依据②的表格回填综测总表（已接入）

用法：
  双击同目录的「启动.command」；或把文件夹拖到它图标上（路径会被记住，
  进入任一任务页后自动带上）。
"""
import os
import sys

from ui.ui_common import (Fonts, default_out_root, open_path,
                       topmost_once, LogPane, setup_style, header_bar,
                       primary_button, secondary_button, hint_bar, card_section,
                       section_label, page_label, muted_label, fade_in,
                       ScrollableFrame, place_window, RoundedCard, RADIUS_CARD,
                       entry, sp,
                       C_PAGE, C_SURFACE, C_BORDER, C_TEXT, C_MUTED, C_ACCENT,
                       C_ACCENT_SOFT, C_ACCENT_TINT, C_ACCENT_HOVER,
                       C_CARD_HOVER, C_BORDER_HOVER,
                       # 动效
                       stagger_highlight)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


class Home(object):

    def __init__(self, root, preset_path=''):
        self.root = root
        self.preset = preset_path
        self.F = Fonts()
        self.out_root = tk.StringVar(value=default_out_root())
        self.status = tk.StringVar(value='就绪')
        self.opened = {}

        root.title('综测核算程序')
        setup_style(root)
        place_window(root, 980, 780)

        self._build()
        fade_in(root)
        topmost_once(root)

    # ---------- 界面 ----------
    def _build(self):
        r = self.root
        pad = {'padx': 14, 'pady': 6}

        header_bar(r, '综测核算程序',
                   '太原科技大学 · 计算机科学与技术学院　｜　先选任务，再进入对应页面',
                   self.F, status_var=self.status, window=self.root,
                   right_builder=lambda rf: secondary_button(
                       rf, '设置', command=self._open_out,
                       padx=12, pady=6).pack(side='right')).pack(fill='x')

        scroll = ScrollableFrame(r)
        scroll.pack(fill='both', expand=True)
        host = scroll.container

        # 输出根目录
        of = card_section(host, '产物输出根目录', F=self.F)
        of.pack(fill='x', **pad)
        fr = ttk.Frame(of)
        fr.pack(fill='x')
        ttk.Label(fr, text='输出根目录：').pack(side='left')
        entry(fr, textvariable=self.out_root, font=self.F.monos).pack(
            side='left', fill='x', expand=True, padx=(6, 8))

        def browse():
            d = filedialog.askdirectory(title='选择产物输出根目录',
                                        initialdir=self.out_root.get())
            if d:
                self.out_root.set(d)

        secondary_button(fr, '浏览…', command=browse, padx=12, pady=5).pack(side='left')
        secondary_button(fr, '打开', padx=12, pady=5, command=lambda: (os.makedirs(
            self.out_root.get(), exist_ok=True), open_path(self.out_root.get()))
            ).pack(side='left', padx=(8, 0))
        muted_label(of, '两个任务的产物默认都落在下载文件夹内；可改成任意位置。',
                    self.F).pack(anchor='w', pady=(6, 0))

        # 任务卡片
        section_label(host, '选择任务', self.F).pack(anchor='w', padx=14, pady=(12, 0))
        cards = tk.Frame(host, bg=C_PAGE)
        cards.pack(fill='both', expand=True, padx=14, pady=4)

        c1 = self._card(cards, '1', '综测两表核对',
                        '核验「综合素质测评详情表」与「综合素质测评表」两份表是否一致。\n'
                        '按 必须处理 / 需确认 / 提示 三级输出问题清单。只读源表，绝不写回。',
                        self._open_verify, enabled=True)
        c2 = self._card(cards, '2', '挂科 / 附加分筛查',
                        '识别必修课 → 判定挂科学生（可整行标黄）→ 认定必修课 100 分附加分资格。\n'
                        '产出《挂科与附加分表.xlsx》（7 个 Sheet），供后续回填综测总表。',
                        self._open_screening, enabled=True)
        c3 = self._card(cards, '3', '回填成绩表',
                        '依据 ② 的产出，把「是否挂科」「附加分」回填到各年级综测总表。\n'
                        '学号+班级双重确认身份，异常逐条报告，写入前先备份。',
                        self._open_fillback, enabled=True)

        # —— 三张卡片依次亮起边框后回落（入场动效）——
        try:
            stagger_highlight([c1, c2, c3], delay_between=80)
        except Exception:
            pass

        if self.preset:
            page_label(host, '已带入路径：' + self.preset, self.F).pack(
                anchor='w', padx=14, pady=(0, 4))
        hint_bar(host, '提示：也可以把文件夹直接拖到程序图标上，路径会被记住　·　'
                       '运行前请关闭所有打开的 Excel', self.F).pack(fill='x')

    def _card(self, parent, num, title, desc, cmd, enabled=True):
        card = RoundedCard(parent, radius=RADIUS_CARD, pad=16)
        card.pack(fill='x', pady=7)
        inner = card.inner
        top = tk.Frame(inner, bg=C_SURFACE)
        top.pack(fill='x')
        badge = tk.Frame(top, bg=C_ACCENT_SOFT, width=sp(40), height=sp(40))
        badge.pack(side='left', padx=(0, 14))
        badge.pack_propagate(False)
        badge_lbl = tk.Label(badge, text=num, bg=C_ACCENT_SOFT, fg=C_ACCENT,
                             font=self.F.h2)
        badge_lbl.place(relx=0.5, rely=0.5, anchor='center')
        body = tk.Frame(top, bg=C_SURFACE)
        body.pack(side='left', fill='both', expand=True)
        lbl_t = tk.Label(body, text=title, bg=C_SURFACE, fg=C_TEXT, font=self.F.h2)
        lbl_t.pack(anchor='w')
        lbl_d = tk.Label(body, text=desc, bg=C_SURFACE, fg=C_MUTED,
                         font=self.F.small, justify='left', wraplength=580)
        lbl_d.pack(anchor='w', pady=(6, 12))
        btn = primary_button(body, '进入 →' if enabled else '规划中',
                             cmd if enabled else None)
        btn.pack(anchor='e')
        if not enabled:
            btn.configure(state='disabled')

        targets = (inner, top, body, lbl_t, lbl_d)

        def _paint(bg, border, badge_bg, badge_fg):
            for w in targets:
                try:
                    w.configure(bg=bg)
                except Exception:
                    pass
            try:
                badge.configure(bg=badge_bg)
                badge_lbl.configure(bg=badge_bg, fg=badge_fg)
            except Exception:
                pass
            card.recolor(fill=bg, border=border)

        if enabled:
            def _on_enter(_e):
                _paint(C_CARD_HOVER, C_ACCENT, C_ACCENT_TINT, C_ACCENT)

            def _on_leave(_e):
                _paint(C_SURFACE, C_BORDER, C_ACCENT_SOFT, C_ACCENT)

            def _on_press(_e):
                # 按下：边框更深；松开会回到 hover 状态（_on_enter 会再触发）
                _paint(C_CARD_HOVER, C_ACCENT_HOVER, C_ACCENT_TINT, C_ACCENT)

            for w in (card, inner, top, body, lbl_t, lbl_d, badge, badge_lbl):
                w.bind('<Enter>', _on_enter, add='+')
                w.bind('<Leave>', _on_leave, add='+')
                w.bind('<ButtonPress-1>', _on_press, add='+')
        return card

    def _open_out(self):
        try:
            d = self.out_root.get()
            os.makedirs(d, exist_ok=True)
            open_path(d)
        except Exception:
            pass

    # ---------- 打开任务页 ----------
    def _toplevel(self, key, title, pref_w, pref_h):
        if key in self.opened:
            try:
                self.opened[key].lift()
                return None
            except Exception:
                pass
        win = tk.Toplevel(self.root)
        win.title(title)
        win.configure(bg=C_PAGE)
        place_window(win, pref_w, pref_h, min_w=860, min_h=560)
        self.opened[key] = win
        return win

    def _open_verify(self):
        win = self._toplevel('verify', '① 综测两表核对', 1040, 800)
        if win is None:
            return
        import task_verify
        task_verify.build(win, self.F, self.out_root.get(), self.preset)
        topmost_once(win)

    def _open_screening(self):
        win = self._toplevel('screening', '② 挂科 / 附加分筛查', 1080, 840)
        if win is None:
            return
        import task_screening
        task_screening.build(win, self.F, self.out_root.get(), self.preset)
        topmost_once(win)

    def _open_fillback(self):
        win = self._toplevel('fillback', '③ 回填成绩表', 1040, 820)
        if win is None:
            return
        import task_fillback
        task_fillback.build(win, self.F, self.out_root.get(), self.preset)
        topmost_once(win)


def selftest():
    """无头自检：构造首页与三个任务页，全部成功返回 0，否则返回 1。

    供打包后自动化验证用（窗口版 exe 无控制台，靠退出码判定）：
        综测核算程序.exe --selftest   → 退出码 0 表示 UI 构建正常
    """
    import traceback
    try:
        root = tk.Tk()
        root.withdraw()
        home = Home(root, '')
        import task_verify
        import task_screening
        import task_fillback
        for mod in (task_verify, task_screening, task_fillback):
            win = tk.Toplevel(root)
            mod.build(win, home.F, home.out_root.get(), '')
            root.update_idletasks()
            win.update_idletasks()
            win.destroy()
        root.update_idletasks()
        root.destroy()
        return 0
    except Exception:
        try:
            base = (os.path.dirname(sys.executable)
                    if getattr(sys, 'frozen', False) else HERE)
            with open(os.path.join(base, 'selftest_error.txt'), 'w',
                      encoding='utf-8') as fh:
                fh.write(traceback.format_exc())
        except Exception:
            pass
        return 1


def main():
    if '--selftest' in sys.argv:
        sys.exit(selftest())
    preset = sys.argv[1] if len(sys.argv) > 1 else ''
    if preset:
        preset = preset.strip().strip('"').strip("'")
        if not os.path.isdir(preset):
            preset = ''
    root = tk.Tk()
    Home(root, preset)
    root.mainloop()


if __name__ == '__main__':
    main()
