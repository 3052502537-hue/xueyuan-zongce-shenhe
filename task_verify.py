# -*- coding: utf-8 -*-
"""
① 综测两表核对
============================================================================
核验每个班的「综合素质测评详情表」与「综合素质测评表」两表一致性。

★ 重要：核验规则全部来自引擎 verify.py（CAP / 各项判定 / 合并共性问题 / 报告格式），
  本页 **不复制也不修改任何规则**，只做调度与展示。

与旧版 run.py 的差异（更精细的部分）：
  - 进入前先「识别」：显示检测到几个班，并**预警缺表的班**（旧版要跑完才知道）
  - 逐班进度：直接调用引擎的 verify_class，每班完成即回显（旧版是一个黑盒跑完）
  - 结果页更完整：汇总数字 + 班级汇总表 + 问题清单表（按级别排序）
  - 可指定学期标识与输出目录；完成后可打开报告 / 在访达中定位
  - 只读源表，绝不写回（引擎本身也只读）
"""
import os
import glob
import time
import queue
import contextlib
from collections import Counter
from datetime import datetime

from ui.ui_common import (LogPane, Runner, QueueWriter, path_row,
                       open_path, reveal_path, load_engine, card_section,
                       hint_bar, fade_in, animate_progress, secondary_button,
                       entry, checkbox, scroll_table, muted_label,
                       page_scaffold, action_bar)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
DETAIL_KEY, EVAL_KEY = '详情', '测评表'

# 界面「问题清单」按需求隐藏的类型（报告仍完整记录）
HIDDEN_TYPES = ('详情表不平',)
# 无论学生是否通过都保留的类型（整班/文件级要紧问题，无姓名，丢了会漏掉整班故障）
KEEP_ALWAYS_TYPES = ('文件缺失', '文件被占用', '人员缺失')


def filter_display_issues(issues):
    """返回用于界面展示的问题清单：隐藏 HIDDEN_TYPES 里的类型。"""
    return [x for x in issues if str(x.get('类型', '')) not in HIDDEN_TYPES]


def select_display_issues(issues, fail_keys):
    """界面问题清单的取舍：只保留「不通过」学生的问题与原因。

    fail_keys 为不通过学生的 (班级, 姓名) 集合；
    无姓名（班级/文件级）与 KEEP_ALWAYS_TYPES 一律保留。
    其余问题不入界面，但报告 xlsx 仍完整记录。
    """
    out = []
    for x in filter_display_issues(issues):
        name = str(x.get('姓名', '') or '')
        typ = str(x.get('类型', ''))
        if name == '' or typ in KEEP_ALWAYS_TYPES:
            out.append(x)
        elif (str(x.get('班级', '')), name) in fail_keys:
            out.append(x)
    return out


# ============================================================
# 目录识别
# ============================================================
def scan_classes(root_dir):
    """返回 (班级列表, 问题列表)。问题=缺详情表或测评表的班。"""
    if not root_dir or not os.path.isdir(root_dir):
        return [], ['路径无效或不是文件夹']
    subs = sorted(d for d in os.listdir(root_dir)
                  if os.path.isdir(os.path.join(root_dir, d))
                  and not d.startswith('.'))
    bad = []
    for s in subs:
        fs = [f for f in glob.glob(os.path.join(root_dir, s, '*.xlsx'))
              if not os.path.basename(f).startswith(('~$', '.~', '._'))]
        has_d = any(DETAIL_KEY in os.path.basename(f) for f in fs)
        has_e = any(EVAL_KEY in os.path.basename(f) for f in fs)
        if not has_d and not has_e:
            bad.append('%s（两份表都缺）' % s)
        elif not has_d:
            bad.append('%s（缺详情表）' % s)
        elif not has_e:
            bad.append('%s（缺测评表）' % s)
    return subs, bad


def report_path(out_dir, term, data):
    os.makedirs(out_dir, exist_ok=True)
    if not term:
        term = os.path.basename(os.path.dirname(os.path.abspath(data))) or '综测'
    stamp = datetime.now().strftime('%Y%m%d-%H%M')
    return os.path.join(out_dir, '综测分数核验报告_%s_%s.xlsx' % (term, stamp))


# ============================================================
# 页面
# ============================================================
def build(win, F, out_root_default, preset=''):
    """在 Toplevel(win) 上构建页面。"""
    try:
        verify, eng_path = load_engine(HERE, 'verify.py', 'zt_verify')
    except Exception as e:
        messagebox.showerror('缺少引擎', str(e))
        return

    var_data = tk.StringVar(value=preset)
    var_term = tk.StringVar()
    var_out = tk.StringVar(value=os.path.join(out_root_default, '核验报告'))
    var_open = tk.BooleanVar(value=True)
    var_reveal = tk.BooleanVar(value=False)
    var_status = tk.StringVar(value='就绪')

    last_out = {'path': ''}
    state = {'classes': []}

    pad = {'padx': 12, 'pady': 5}

    host, var_big, var_side = page_scaffold(
        win, F, '① 综测两表核对',
        '比对「综合素质测评详情表」与「综合素质测评表」，'
        '按 必须处理 / 需确认 / 提示 三级列出问题。源表只读，绝不写回。',
        '检测到的班级', var_status)

    # ---- 数据来源
    f1 = card_section(host, '数据来源', F=F)
    f1.pack(fill='x', **pad)
    lbl_scan = ttk.Label(f1, text='年级文件夹：其下每个子文件夹为一个班，班内含两份 xlsx。',
                         font=F.small)

    def on_change():
        subs, bad = scan_classes(var_data.get().strip())
        state['classes'] = subs
        var_big.set(str(len(subs)))
        if not subs:
            lbl_scan.configure(text='未检测到班级子文件夹（或路径无效）。')
            return
        txt = '检测到 %d 个班。' % len(subs)
        if bad:
            txt += '  ⚠ 缺表 %d 个：' % len(bad) + '、'.join(bad[:6])
            if len(bad) > 6:
                txt += ' 等'
        else:
            txt += '  两份表齐全。'
        lbl_scan.configure(text=txt)

    path_row(f1, '年级文件夹', var_data, on_change, '选择年级文件夹（其下每班一个子文件夹）')
    secondary_button(f1, '识别', command=on_change, padx=12, pady=5).pack(
        anchor='w', pady=(2, 0))
    lbl_scan.pack(anchor='w')

    # ---- 选项
    f2 = card_section(host, '选项', F=F)
    f2.pack(fill='x', **pad)
    r1 = ttk.Frame(f2)
    r1.pack(fill='x', pady=2)
    ttk.Label(r1, text='学期标识：', width=10).pack(side='left')
    entry(r1, textvariable=var_term, width=22).pack(side='left')
    ttk.Label(r1, text='留空则自动取上级目录名', font=F.small).pack(side='left', padx=8)

    def browse_out():
        d = filedialog.askdirectory(title='选择报告输出目录',
                                    initialdir=var_out.get() or out_root_default)
        if d:
            var_out.set(d)

    r2 = ttk.Frame(f2)
    r2.pack(fill='x', pady=2)
    ttk.Label(r2, text='输出目录：', width=10).pack(side='left')
    entry(r2, textvariable=var_out).pack(side='left', fill='x', expand=True, padx=(0, 6))
    secondary_button(r2, '浏览…', command=browse_out, padx=12, pady=5).pack(side='left')

    r3 = ttk.Frame(f2)
    r3.pack(fill='x', pady=2)
    checkbox(r3, text='完成后打开报告', variable=var_open, F=F).pack(side='left')
    checkbox(r3, text='完成后在访达中定位报告', variable=var_reveal,
             F=F).pack(side='left', padx=16)

    # ---- 执行
    f3 = card_section(host, '执行', F=F)
    f3.pack(fill='x', **pad)
    _prev, btn_run, btn_open, btn_reveal, pb = action_bar(
        f3, '▶ 开始核验', lambda: logpane.clear(),
        open_text='打开报告', progress='determinate')

    # ---- 结果
    f4 = card_section(host, '结果', F=F)
    f4.pack(fill='both', expand=False, **pad)
    sumf = ttk.Frame(f4)
    sumf.pack(fill='x')
    lbl_sum = ttk.Label(sumf, text='尚未核验', font=F.sec)
    lbl_sum.pack(anchor='w')

    nb = ttk.Notebook(f4)
    nb.pack(fill='both', expand=True, pady=(6, 0))

    tab_cls = ttk.Frame(nb)
    nb.add(tab_cls, text='班级汇总')
    cols_c = ('班级', '学生数', '通过', '不通过', '问题条数', '仅详情', '仅测评')
    wrap_c, tree_cls = scroll_table(tab_cls, cols_c, height=10)
    wrap_c.pack(fill='both', expand=True, pady=(6, 0))
    for c in cols_c:
        tree_cls.heading(c, text=c)
        tree_cls.column(c, width=110 if c == '班级' else 80, anchor='center')

    tab_iss = ttk.Frame(nb)
    nb.add(tab_iss, text='问题清单')
    muted_label(tab_iss, '仅列出「不通过」学生的问题与原因（含文件/班级级要紧问题）；'
                         '其余提醒类问题不在此列，完整清单见核验报告。', F).pack(
        anchor='w', pady=(6, 0))
    cols_i = ('级别', '班级', '姓名', '部位', '类型', '说明', '处理建议')
    wrap_i, tree_iss = scroll_table(tab_iss, cols_i, height=14)
    wrap_i.pack(fill='both', expand=True, pady=(6, 0))
    w = {'级别': 80, '班级': 110, '姓名': 80, '部位': 100, '类型': 120,
         '说明': 420, '处理建议': 200}
    for c in cols_i:
        tree_iss.heading(c, text=c)
        tree_iss.column(c, width=w[c], anchor='w')

    # ---- 日志
    f5 = card_section(host, '日志', F=F)
    f5.pack(fill='both', expand=True, **pad)
    logpane = LogPane(f5, height=12)
    logpane.pack(fill='both', expand=True)

    # ============================================================
    # 后台任务
    # ============================================================
    def handler(kind, payload):
        if kind == 'status':
            var_status.set(payload)
        elif kind == 'progress':
            animate_progress(pb, payload)
        elif kind == 'summary':
            lbl_sum.configure(text=payload)
            parts = payload.split('　｜　')
            var_side.set('\n'.join(parts[:2]))
        elif kind == 'tree_cls':
            for it in tree_cls.get_children():
                tree_cls.delete(it)
            for s in payload:
                tree_cls.insert('', 'end', values=(
                    s.get('班级', ''), s.get('学生数', 0), s.get('通过', 0),
                    s.get('不通过', 0), s.get('问题条数', 0),
                    s.get('仅详情', 0), s.get('仅测评', 0)))
        elif kind == 'tree_iss':
            for it in tree_iss.get_children():
                tree_iss.delete(it)
            for r in payload:
                desc = str(r.get('说明', ''))
                if len(desc) > 120:
                    desc = desc[:120] + '…'
                tree_iss.insert('', 'end', values=(
                    r.get('级别', ''), r.get('班级', ''), r.get('姓名', ''),
                    r.get('部位', ''), r.get('类型', ''), desc,
                    r.get('处理建议', '')))
        elif kind == 'done_out':
            last_out['path'] = payload
            btn_open.configure(state='normal')
            btn_reveal.configure(state='normal')

    def on_finish():
        btn_run.configure(state='normal')
        pb.stop()
        if var_open.get() and last_out['path']:
            open_path(last_out['path'])
        if var_reveal.get() and last_out['path']:
            reveal_path(last_out['path'])

    runner = Runner(win, logpane, handler=handler, on_finish=on_finish)

    def worker(q, cfg):
        # ⚠ 注意：子线程里绝不能碰 tk.StringVar / 任何 Tk 控件（会抛
        #   "main thread is not in main loop"）。所有界面取值都在 start() 里
        #   快照成普通 dict 后传进来。
        data = cfg['data']
        if not data or not os.path.isdir(data):
            raise RuntimeError('请先选择有效的年级文件夹')
        classes = sorted(d for d in os.listdir(data)
                         if os.path.isdir(os.path.join(data, d))
                         and not d.startswith('.'))
        if not classes:
            raise RuntimeError('该目录下没有班级子文件夹')

        q.put(('status', '核验中…'))
        q.put(('log', '引擎：%s' % os.path.basename(eng_path)))
        q.put(('log', '共 %d 个班，开始逐班核验（源表只读）' % len(classes)))

        all_rows, issues_all, summary_classes = [], [], []
        t0 = time.time()
        for i, cls in enumerate(classes, 1):
            q.put(('progress', i * 100.0 / len(classes)))
            q.put(('log', '  [%d/%d] %s' % (i, len(classes), cls)))
            with contextlib.redirect_stdout(QueueWriter(q)):
                rows, issues, summary = verify.verify_class(
                    cls, os.path.join(data, cls))
            all_rows.extend(rows)
            issues_all.extend(issues)
            if summary:
                summary_classes.append(summary)

        # === 以下三步与原 run_verification 完全一致（合并共性问题 + 重算问题条数）
        q.put(('log', '合并整班共性问题…'))
        issues_all = verify.merge_common_issues(issues_all)
        cnt = Counter(r['班级'] for r in issues_all)
        for s in summary_classes:
            s['问题条数'] = cnt.get(s['班级'], 0)

        total = len(all_rows)
        fail = sum(1 for x in all_rows if x.get('结果') == '不通过')
        lv = Counter(x.get('级别') for x in issues_all)

        out = report_path(cfg['out'], cfg['term'], data)
        q.put(('status', '生成报告…'))
        with contextlib.redirect_stdout(QueueWriter(q)):
            verify.build_report(out, all_rows, issues_all, summary_classes)

        q.put(('tree_cls', summary_classes))
        lv_order = {'必须处理': 0, '需确认': 1, '提示': 2}
        issues_sorted = sorted(
            issues_all, key=lambda x: (lv_order.get(x.get('级别', '提示'), 3),
                                       x.get('班级', ''), x.get('姓名', '')))
        # 界面问题清单：只保留「不通过」学生的问题与原因（其余照常完整写入报告）
        fail_keys = set()
        for r in all_rows:
            if r.get('结果') == '不通过':
                fail_keys.add((str(r.get('班级', '')), str(r.get('姓名', ''))))
        issues_show = select_display_issues(issues_sorted, fail_keys)
        hidden_n = len(issues_sorted) - len(issues_show)
        q.put(('tree_iss', issues_show[:1500]))
        if hidden_n:
            q.put(('log', '（问题清单仅显示不通过学生的问题：不通过 %d 人，'
                          '隐藏其余 %d 条；完整清单见报告）'
                   % (len(fail_keys), hidden_n)))
        if len(issues_show) > 1500:
            q.put(('log', '（问题清单共 %d 条，界面展示前 1500 条，完整版见报告）'
                   % len(issues_show)))
        q.put(('summary', '总人数 %d　通过 %d　不通过 %d　｜　'
                          '必须处理 %d　需确认 %d　提示 %d　｜　用时 %.1fs'
               % (total, total - fail, fail,
                  lv.get('必须处理', 0), lv.get('需确认', 0), lv.get('提示', 0),
                  time.time() - t0)))
        q.put(('done_out', out))
        q.put(('log', '✔ 核验完成：%d 人，通过 %d，不通过 %d' % (total, total - fail, fail)))
        q.put(('log', '报告：' + out))
        q.put(('status', '完成'))

    def start():
        if runner.busy:
            return
        data = var_data.get().strip().strip('"').strip("'")
        if not data or not os.path.isdir(data):
            messagebox.showwarning('提示', '请先选择有效的年级文件夹')
            return
        logpane.clear()
        for it in tree_cls.get_children():
            tree_cls.delete(it)
        for it in tree_iss.get_children():
            tree_iss.delete(it)
        pb['value'] = 0
        btn_run.configure(state='disabled')
        btn_open.configure(state='disabled')
        btn_reveal.configure(state='disabled')
        # 在主线程把界面取值快照成普通 dict，再交给子线程
        cfg = {'data': data,
               'term': var_term.get().strip(),
               'out': var_out.get().strip()}
        runner.run(lambda q: worker(q, cfg))

    btn_run.configure(command=start)
    btn_open.configure(command=lambda: open_path(last_out['path'])
                       if last_out['path'] and os.path.exists(last_out['path'])
                       else messagebox.showinfo('提示', '还没有报告'))
    btn_reveal.configure(command=lambda: reveal_path(last_out['path'])
                         if last_out['path'] and os.path.exists(last_out['path'])
                         else messagebox.showinfo('提示', '还没有报告'))

    if preset and os.path.isdir(preset):
        win.after(200, on_change)

    hint_bar(host, '源表只读：本页绝不写回原文件，报告为独立新文件。', F).pack(
        anchor='w', fill='x')
    fade_in(win)

    return win
