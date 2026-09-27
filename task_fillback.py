# -*- coding: utf-8 -*-
"""
③ 回填成绩表
============================================================================
依据 ② 产出的《挂科与附加分表.xlsx》，把「是否挂科」「附加分」回填到
各年级「综测成绩表」总表。判定与写入规则全部在引擎 回填引擎.py 里，
本页只做调度与展示。

规则速记：
  是否挂科 ← 「挂科汇总」学号集合：命中「是」，否则「否」
  附加分   ← 「附加分汇总」的「100分门数」：1 门=1，无则 0
  身份确认  ：学号 + 班级双重确认；学号位数异常、班级不一致都会报告
  年级对齐  ：源表含多年级时，按目录学生学号自动锁定年级，避免误报缺人
  写入前    ：一律先备份；一律按名单覆盖
"""
import os
import contextlib

from ui.ui_common import (LogPane, Runner, QueueWriter, path_row, open_path,
                       reveal_path, load_engine, card_section, hint_bar,
                       fade_in, secondary_button, entry, checkbox, scroll_table,
                       page_scaffold, action_bar)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_NAME = '挂科与附加分表.xlsx'


def auto_find_src(roots):
    """在给定根目录里递归找最新的《挂科与附加分表.xlsx》。"""
    found = []
    for r in roots:
        if not r or not os.path.isdir(r):
            continue
        for root, dirs, fs in os.walk(r):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for f in fs:
                if f == SRC_NAME and not f.startswith('~$'):
                    p = os.path.join(root, f)
                    try:
                        found.append((os.path.getmtime(p), p))
                    except OSError:
                        pass
    found.sort()
    return found[-1][1] if found else ''


def build(win, F, out_root_default, preset=''):
    try:
        eng, eng_path = load_engine(HERE, '回填引擎.py', 'fb_engine')
    except Exception as e:
        messagebox.showerror('缺少引擎', str(e))
        return

    var_tables = tk.StringVar(value=preset or '')
    var_src = tk.StringVar(value='')
    var_out = tk.StringVar(value=out_root_default)
    var_backup = tk.BooleanVar(value=True)
    var_open = tk.BooleanVar(value=True)
    var_status = tk.StringVar(value='就绪')
    last_out = {'path': ''}
    pad = {'padx': 12, 'pady': 5}

    host, var_big, var_side = page_scaffold(
        win, F, '③ 回填成绩表',
        '依据 ② 的产出回填综测总表：是否挂科 = 命中名单→是/否则→否；'
        '附加分 = 100 分门数（无则 0）',
        '检测到的综测总表', var_status)

    # ---------- 数据来源 ----------
    f1 = card_section(host, '数据来源', F=F)
    f1.pack(fill='x', **pad)

    fr_src = ttk.Frame(f1)
    fr_src.pack(fill='x', pady=2)
    ttk.Label(fr_src, text='源表', width=10).pack(side='left')
    entry(fr_src, textvariable=var_src).pack(side='left', fill='x',
                                             expand=True, padx=6)

    def browse_src():
        init = os.path.dirname(var_src.get()) or out_root_default
        p = filedialog.askopenfilename(title='选择《%s》' % SRC_NAME,
                                       filetypes=[('Excel', '*.xlsx')],
                                       initialdir=init if os.path.isdir(init) else None)
        if p:
            var_src.set(p)
            on_src_change()

    secondary_button(fr_src, '浏览…', command=browse_src, padx=12, pady=5).pack(side='left')

    def do_auto():
        p = auto_find_src([out_root_default, preset, var_tables.get()])
        if p:
            var_src.set(p)
            on_src_change()
        else:
            lbl_src.configure(text='未自动找到《%s》，请手动选择，或先跑 ② 生成。'
                                   % SRC_NAME)

    secondary_button(fr_src, '自动查找', command=do_auto, padx=12, pady=5).pack(
        side='left', padx=(6, 0))
    lbl_src = ttk.Label(f1, text='', font=F.small, wraplength=780, justify='left')
    lbl_src.pack(anchor='w')

    def on_src_change(*_a):
        p = var_src.get().strip().strip('"').strip("'")
        if not p or not os.path.isfile(p):
            lbl_src.configure(text='源表路径无效。')
            return
        try:
            _f, _a, info = eng.load_source(p)
            lbl_src.configure(text='源表就绪：挂科 %d 人 · 附加分 %d 人/%d 门次 · '
                                   '涉及年级 %s'
                              % (info['fail'], info['add'], info['add_cnt'],
                                 '、'.join(info['grades']) or '（无年级列）'))
        except Exception as e:
            lbl_src.configure(text='读取失败：%s' % e)

    def on_tables_change(*_a):
        d = var_tables.get().strip().strip('"').strip("'")
        if not d or not os.path.isdir(d):
            lbl_tab.configure(text='路径无效或不是文件夹。')
            var_big.set('0')
            return
        try:
            targets = eng.collect_targets(d)
        except Exception as e:
            lbl_tab.configure(text='读取失败：%s' % e)
            return
        if not targets:
            lbl_tab.configure(text='该目录下没有找到综测总表'
                                   '（需含「姓名 / 学号 / 是否挂科」表头）。')
            var_big.set('0')
            return
        n = sum(len(f) for _l, f in targets)
        var_big.set(str(n))
        lbl_tab.configure(text='检测到 %d 个年级目录、%d 个综测总表：%s'
                          % (len(targets), n,
                             '  '.join('%s(%d)' % (l, len(f))
                                       for l, f in targets)))

    path_row(f1, '总表目录', var_tables, on_tables_change, '选择综测成绩表所在目录')
    secondary_button(f1, '识别', command=on_tables_change, padx=12, pady=5).pack(
        anchor='w', pady=(2, 0))
    lbl_tab = ttk.Label(f1, text='可以是年级目录（里面直接是若干专业总表），'
                                 '也可以是含多个年级子目录的父目录。',
                        font=F.small, wraplength=780, justify='left')
    lbl_tab.pack(anchor='w')

    # ---------- 选项 ----------
    f2 = card_section(host, '选项', F=F)
    f2.pack(fill='x', **pad)
    checkbox(f2, text='写入前先备份原表（推荐）', variable=var_backup,
             F=F).pack(anchor='w', pady=3)
    checkbox(f2, text='完成后自动打开报告', variable=var_open,
             F=F).pack(anchor='w', pady=3)

    def browse_out():
        d = filedialog.askdirectory(title='选择报告输出目录',
                                    initialdir=var_out.get() or out_root_default)
        if d:
            var_out.set(d)

    ro = ttk.Frame(f2)
    ro.pack(fill='x', pady=(6, 0))
    ttk.Label(ro, text='报告输出：', width=10).pack(side='left')
    entry(ro, textvariable=var_out).pack(side='left', fill='x', expand=True, padx=(0, 6))
    secondary_button(ro, '浏览…', command=browse_out, padx=12, pady=5).pack(side='left')

    # ---------- 执行 ----------
    f3 = card_section(host, '执行', F=F)
    f3.pack(fill='x', **pad)
    btn_prev, btn_run, btn_open, btn_rev, pb = action_bar(
        f3, '▶ 开始回填', lambda: logpane.clear(),
        preview_text='🔍 预览（不改任何文件）', open_text='打开报告')

    # ---------- 各表一览 ----------
    f4 = card_section(host, '各表一览', F=F)
    f4.pack(fill='both', expand=False, **pad)
    cols = ('目录', '文件', '学生行', '挂科(是)', '否', '附加分人数', '附加分总分')
    wrap, tree = scroll_table(f4, cols, height=11)
    wrap.pack(fill='both', expand=True, pady=(6, 0))
    w = {'目录': 130, '文件': 200, '学生行': 70, '挂科(是)': 80,
         '否': 70, '附加分人数': 90, '附加分总分': 90}
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=w[c], anchor='center')

    # ---------- 日志 ----------
    f5 = card_section(host, '日志', F=F)
    f5.pack(fill='both', expand=True, **pad)
    logpane = LogPane(f5, height=12)
    logpane.pack(fill='both', expand=True)

    # ============================================================
    def handler(kind, payload):
        if kind == 'status':
            var_status.set(payload)
        elif kind == 'tree':
            for it in tree.get_children():
                tree.delete(it)
            for s in payload:
                tree.insert('', 'end', values=(s['dir'], s['file'], s['rows'],
                                               s['yes'], s['no'], s['add'],
                                               s['add_cnt']))
        elif kind == 'summary':
            logpane.write(payload, 'ok')
            var_side.set(str(payload).strip('─ ').replace(' · ', '\n'))
        elif kind == 'done_out':
            last_out['path'] = payload
            btn_open.configure(state='normal')
            btn_rev.configure(state='normal')

    def on_finish():
        btn_prev.configure(state='normal')
        btn_run.configure(state='normal')
        pb.stop()
        if var_open.get() and last_out['path']:
            open_path(last_out['path'])

    runner = Runner(win, logpane, handler=handler, on_finish=on_finish)

    def worker(q, cfg):
        # ⚠ 子线程不能碰 tk 变量，所有取值都在 do_* 里快照成 dict 传进来
        if not cfg['src'] or not os.path.isfile(cfg['src']):
            raise RuntimeError('请先选择源表《%s》' % SRC_NAME)
        if not cfg['tables'] or not os.path.isdir(cfg['tables']):
            raise RuntimeError('请先选择总表目录')

        q.put(('status', '读取源表…'))
        src, targets, plans = eng.prepare(cfg['tables'], cfg['src'])
        q.put(('log', '源表：%s' % src[2]['path']))
        for p in plans:
            q.put(('log', '  %s → 目标年级 %s（%s）'
                   % (p['label'], p['grade'] or '未识别', p['note'])))

        q.put(('status', '空跑统计…'))
        with contextlib.redirect_stdout(QueueWriter(q)):
            res = eng.phase_dryrun(plans, src)
            eng.print_issues(res['issues'])
        q.put(('tree', res['per_file']))
        t = res['total']
        q.put(('summary', '── 合计：%d 行 · 是否挂科 是=%d 否=%d · '
                          '附加分 %d 人 / %d 门次 ──'
               % (t['rows'], t['yes'], t['no'], t['add'], t['add_cnt'])))
        n_iss = len([i for i in res['issues'] if i['kind'] != '跨年级跟班生'])
        if n_iss:
            q.put(('summary', '注意：有 %d 条需确认的异常，详见日志与报告' % n_iss))

        if not cfg['write']:
            q.put(('log', '✔ 预览完成，未修改任何文件'))
            q.put(('status', '预览完成（未修改任何文件）'))
            return

        q.put(('status', '备份 + 写入 + 核验…'))
        with contextlib.redirect_stdout(QueueWriter(q)):
            ok, backups = eng.phase_apply(plans, src, do_backup=cfg['backup'])
            rep = eng.write_report(cfg['out'], plans, src, res, backups, ok)
        q.put(('done_out', rep))
        q.put(('log', '✔ 报告：' + rep))
        q.put(('status', '完成' if ok else '核验未通过，请检查日志'))

    def snapshot(write):
        return {'src': var_src.get().strip().strip('"').strip("'"),
                'tables': var_tables.get().strip().strip('"').strip("'"),
                'out': var_out.get().strip() or out_root_default,
                'backup': var_backup.get(),
                'write': write}

    def guard():
        s = snapshot(False)
        if not s['src'] or not os.path.isfile(s['src']):
            messagebox.showwarning('提示', '请先选择源表《%s》（可点「自动查找」）' % SRC_NAME)
            return None
        if not s['tables'] or not os.path.isdir(s['tables']):
            messagebox.showwarning('提示', '请先选择总表目录')
            return None
        return s

    def do_preview():
        s = guard()
        if s is None:
            return
        s['write'] = False
        btn_prev.configure(state='disabled')
        btn_run.configure(state='disabled')
        pb.start(12)
        runner.run(lambda q: worker(q, s))

    def do_run():
        s = guard()
        if s is None:
            return
        s['write'] = True
        if not messagebox.askyesno(
                '确认回填',
                '将按名单覆盖「是否挂科」「附加分」两列：\n\n'
                '  总表目录：%s\n  源表：%s\n\n'
                '%s\n继续？'
                % (s['tables'], os.path.basename(s['src']),
                   '写入前会先备份原表。' if s['backup']
                   else '⚠ 本次未勾选备份，原表将被直接覆盖！')):
            return
        btn_prev.configure(state='disabled')
        btn_run.configure(state='disabled')
        pb.start(12)
        runner.run(lambda q: worker(q, s))

    btn_prev.configure(command=do_preview)
    btn_run.configure(command=do_run)
    btn_open.configure(command=lambda: open_path(last_out['path']))
    btn_rev.configure(command=lambda: reveal_path(last_out['path']))

    # ---------- 规则速查 ----------
    f6 = card_section(host, '字段映射 / 已知坑', F=F)
    f6.pack(fill='x', **pad)
    for s in ['综测总表 A=姓名、B=学号；E(5)=附加分、K(11)=是否挂科（按表头文字定位）',
              'E 列参与 F 列公式 =D+E，未获资格也必须填 0，不能留空',
              '学号带 Excel 文本前缀撇号会自动归一化；非 12 位会逐条报告',
              '跨年级跟班生（学号年级 ≠ 班级年级）属正常，按学号匹配，不判错',
              '总表比名单少人属总表缺人，降级为警告，不算填充失败']:
        ttk.Label(f6, text='· ' + s, font=F.small, wraplength=800,
                  justify='left').pack(anchor='w', pady=1)

    # ---------- 启动即自动带入 ----------
    if var_tables.get():
        on_tables_change()
    do_auto()

    hint_bar(host, '回填前会自动备份原表；「预览」不修改任何文件，可先核对。', F).pack(
        anchor='w', fill='x')
    fade_in(win)
    return win
