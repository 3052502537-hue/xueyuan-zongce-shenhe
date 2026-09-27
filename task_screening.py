# -*- coding: utf-8 -*-
"""
② 挂科 / 附加分筛查
============================================================================
识别必修课 → 判定挂科学生（可整行标黄）→ 认定必修课 100 分附加分资格
→ 产出《挂科与附加分表.xlsx》（供 ③ 回填综测总表使用）

判定规则全部在引擎 挂科引擎.py 里，本页只做调度与展示。

规则速记：
  必修 = 课程列有成绩人数 ≥ 班级实际人数 90%（零星重修/公选列不计）
  挂科 = 必修列 <60 / 「不及格」/ 缺考·缓考·取消资格等未知文字
  附加分 = 必修列**数字**成绩 == 100（五档「优」不算；只列资格，不算分值）
  注意：「人」与「门次」不是一回事，一人可挂多门。
"""
import os
import glob
import time
import contextlib

from ui.ui_common import (LogPane, Runner, QueueWriter, path_row, open_path,
                       reveal_path, load_engine, card_section, hint_bar,
                       fade_in, secondary_button, entry, checkbox, scroll_table,
                       page_scaffold, action_bar)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))


def count_xlsx(d):
    if not d or not os.path.isdir(d):
        return 0
    n = 0
    for root, dirs, fs in os.walk(d):
        dirs[:] = [x for x in dirs if not x.startswith('.')]
        for f in fs:
            if (f.lower().endswith('.xlsx')
                    and not f.startswith(('~$', '.~', '._'))
                    and f != '挂科与附加分表.xlsx'):
                n += 1
    return n


def build(win, F, out_root_default, preset=''):
    try:
        eng, eng_path = load_engine(HERE, '挂科引擎.py', 'gk_engine')
    except Exception as e:
        messagebox.showerror('缺少引擎', str(e))
        return

    var_data = tk.StringVar(value=preset)
    var_out = tk.StringVar(value=out_root_default)
    var_hl = tk.BooleanVar(value=True)
    var_vf = tk.BooleanVar(value=True)
    var_xl = tk.BooleanVar(value=True)
    var_au = tk.BooleanVar(value=True)
    var_open = tk.BooleanVar(value=True)
    var_status = tk.StringVar(value='就绪')

    last_out = {'path': ''}
    pad = {'padx': 12, 'pady': 5}

    host, var_big, var_side = page_scaffold(
        win, F, '② 挂科 / 附加分筛查',
        '必修 = 有成绩人数 ≥ 90%　｜　挂科 = 必修列 <60 / 不及格 / 缺考等　｜　'
        '附加分 = 必修列数字 100 分（五档「优」不算）',
        '检测到的班级成绩表', var_status)

    # ---- 数据来源
    f1 = card_section(host, '数据来源', F=F)
    f1.pack(fill='x', **pad)
    lbl_scan = ttk.Label(f1, text='可以是学期根目录（含多个年级子目录），也可直接选某个年级文件夹。',
                         font=F.small)

    def on_change():
        d = var_data.get().strip()
        if not os.path.isdir(d):
            lbl_scan.configure(text='路径无效或不是文件夹。')
            var_big.set('0')
            return
        try:
            files = eng.collect_xlsx(d)
        except Exception as e:
            lbl_scan.configure(text='读取失败：%s' % e)
            return
        var_big.set(str(len(files)))
        if not files:
            lbl_scan.configure(text='该目录下没有找到班级成绩表（.xlsx）。')
            return
        g = {}
        for f in files:
            k = os.path.basename(os.path.dirname(f))
            g[k] = g.get(k, 0) + 1
        lbl_scan.configure(text='检测到 %d 个班级成绩表：' % len(files)
                           + '  '.join('%s(%d)' % (k, v) for k, v in sorted(g.items())))

    path_row(f1, '成绩文件夹', var_data, on_change, '选择成绩文件夹')
    secondary_button(f1, '识别', command=on_change, padx=12, pady=5).pack(
        anchor='w', pady=(2, 0))
    lbl_scan.pack(anchor='w')

    # ---- 选项
    f2 = card_section(host, '选项', F=F)
    f2.pack(fill='x', **pad)

    def opt(parent, var, text, desc):
        fr = ttk.Frame(parent)
        fr.pack(fill='x', pady=3)
        cb = checkbox(fr, text=text, variable=var, command=sync_opt, F=F)
        cb.pack(side='left')
        ttk.Label(fr, text=desc, font=F.small).pack(side='left', padx=(10, 0))
        return cb

    # ⚠ 必须先定义 sync_opt，再 opt(...)——闭包在调用时才解析名字
    def sync_opt():
        cb_vf.configure(state='disabled' if not var_hl.get() else 'normal')
        cb_au.configure(state='disabled' if not var_xl.get() else 'normal')

    cb_hl = opt(f2, var_hl, '写入挂科标黄到源文件',
                '挂科学生整行标黄 FFFFFF00（会直接修改原表）')
    cb_vf = opt(f2, var_vf, '写入后做交叉核验',
                '独立重算并与实际标黄逐行比对（多标/漏标/残缺）')
    cb_xl = opt(f2, var_xl, '生成《挂科与附加分表.xlsx》',
                '产出结构化表格，供程序 ③ 回填综测总表')
    cb_au = opt(f2, var_au, '包含「课程认定明细」Sheet',
                '审计每列为何算必修/选修（行数最多，可关）')
    opt(f2, var_open, '完成后自动打开结果', '')
    sync_opt()

    def browse_out():
        d = filedialog.askdirectory(title='选择输出目录',
                                    initialdir=var_out.get() or out_root_default)
        if d:
            var_out.set(d)

    ro = ttk.Frame(f2)
    ro.pack(fill='x', pady=(6, 0))
    ttk.Label(ro, text='输出目录：', width=10).pack(side='left')
    entry(ro, textvariable=var_out).pack(side='left', fill='x', expand=True, padx=(0, 6))
    secondary_button(ro, '浏览…', command=browse_out, padx=12, pady=5).pack(side='left')
    sync_opt()

    # ---- 执行
    f3 = card_section(host, '执行', F=F)
    f3.pack(fill='x', **pad)
    btn_prev, btn_run, btn_open, btn_rev, pb = action_bar(
        f3, '▶ 开始处理', lambda: logpane.clear(),
        preview_text='🔍 预览（不改任何文件）', open_text='打开结果')

    # ---- 各班一览
    f4 = card_section(host, '各班一览', F=F)
    f4.pack(fill='both', expand=False, **pad)
    cols = ('年级', '班级', '在册', '必修', '挂科人数', '挂科门次', '附加分人数', '100分门次')
    wrap, tree = scroll_table(f4, cols, height=11)
    wrap.pack(fill='both', expand=True, pady=(6, 0))
    w = {'年级': 70, '班级': 150, '在册': 60, '必修': 60,
         '挂科人数': 80, '挂科门次': 80, '附加分人数': 90, '100分门次': 90}
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=w[c], anchor='center')

    # ---- 日志
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
            for d in payload:
                tree.insert('', 'end', values=(
                    d['grade'] + '级', d['klass'], d['n'], len(d['required']),
                    len(d['fails']), sum(len(x[2]) for x in d['fails']),
                    len(d['perfect']), sum(len(x[2]) for x in d['perfect'])))
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
        # ⚠ 子线程不能碰 tk 变量/控件（会抛 "main thread is not in main loop"），
        #   所有界面取值都在 do_preview / do_run 里快照成 dict 传进来。
        data = cfg['data']
        write = cfg['write']
        if not data or not os.path.isdir(data):
            raise RuntimeError('请先选择有效的成绩文件夹')
        q.put(('status', '扫描…'))
        files = eng.collect_xlsx(data)
        if not files:
            raise RuntimeError('该目录下没有找到班级成绩表（.xlsx）')
        q.put(('log', '找到 %d 个班级成绩表' % len(files)))
        q.put(('status', '空跑判定…'))
        with contextlib.redirect_stdout(QueueWriter(q)):
            results = eng.phase_dryrun(files)
        q.put(('tree', results))

        ts = sum(d['n'] for d in results)
        tf = sum(len(d['fails']) for d in results)
        tfc = sum(len(x[2]) for d in results for x in d['fails'])
        tb = sum(len(d['perfect']) for d in results)
        tbc = sum(len(x[2]) for d in results for x in d['perfect'])
        q.put(('summary', '── 合计：%d 个班 · %d 名学生 · 挂科 %d 人/%d 门次 · '
                          '附加分 %d 人/%d 门次 ──' % (len(results), ts, tf, tfc, tb, tbc)))

        if not write:
            q.put(('status', '预览完成（未修改任何文件）'))
            return

        if cfg['highlight']:
            q.put(('status', '写入标黄…'))
            with contextlib.redirect_stdout(QueueWriter(q)):
                eng.phase_apply(results)
            if cfg['verify']:
                q.put(('status', '交叉核验…'))
                with contextlib.redirect_stdout(QueueWriter(q)):
                    ok = eng.phase_verify(results)
                q.put(('summary', '交叉核验：%s' % (
                    '%d/%d 全部一致' % (len(results), len(results)) if ok
                    else '未通过，请检查')))
            else:
                q.put(('summary', '交叉核验：已跳过'))

        if cfg['xlsx']:
            q.put(('status', '生成表格…'))
            out_dir = cfg['out'] or (data if os.path.isdir(data)
                                     else os.path.dirname(data))
            os.makedirs(out_dir, exist_ok=True)
            out = os.path.join(out_dir, eng.OUT_NAME)
            with contextlib.redirect_stdout(QueueWriter(q)):
                stat = eng.build_workbook(
                    results, out,
                    {'time': time.strftime('%Y-%m-%d %H:%M'),
                     'target': data,
                     'highlight': '是（标准黄 FFFFFF00 整行）' if cfg['highlight'] else '否',
                     'verify': '见日志'},
                    include_course_audit=cfg['audit'])
            q.put(('done_out', out))
            q.put(('log', '✔ 已生成：' + out))
            q.put(('summary', '挂科 %d 人 / %d 门次 · 附加分 %d 人 / %d 门次'
                   % (stat['fail_stu'], stat['fail_cnt'],
                      stat['bonus_stu'], stat['bonus_cnt'])))
        q.put(('status', '完成'))

    def guard():
        d = var_data.get().strip().strip('"').strip("'")
        if not d or not os.path.isdir(d):
            messagebox.showwarning('提示', '请先选择有效的成绩文件夹')
            return False
        if not var_hl.get() and not var_xl.get():
            messagebox.showwarning('提示', '「写入标黄」和「生成表格」都没勾选，不会有任何产出。')
            return False
        return True

    def snapshot(write):
        """主线程里把界面取值快照成普通 dict（子线程不能用 Tk 变量）。"""
        return {'data': var_data.get().strip().strip('"').strip("'"),
                'out': var_out.get().strip(),
                'highlight': var_hl.get(),
                'verify': var_vf.get(),
                'xlsx': var_xl.get(),
                'audit': var_au.get(),
                'write': write}

    def do_preview():
        if runner.busy or not guard():
            return
        btn_prev.configure(state='disabled')
        btn_run.configure(state='disabled')
        pb.start(12)
        var_status.set('预览中…')
        cfg = snapshot(False)   # ⚠ 必须在主线程求值
        runner.run(lambda q: worker(q, cfg))

    def do_run():
        if runner.busy or not guard():
            return
        if var_hl.get():
            if not messagebox.askyesno(
                    '确认写入',
                    '即将：\n· 挂科学生整行标黄（直接修改原 xlsx，不自动备份）\n'
                    '· 生成《挂科与附加分表.xlsx》\n\n确定继续吗？'):
                return
        btn_prev.configure(state='disabled')
        btn_run.configure(state='disabled')
        pb.start(12)
        var_status.set('处理中…')
        cfg = snapshot(True)   # ⚠ 必须在主线程求值
        runner.run(lambda q: worker(q, cfg))

    btn_prev.configure(command=do_preview)
    btn_run.configure(command=do_run)
    btn_open.configure(command=lambda: open_path(last_out['path'])
                       if last_out['path'] and os.path.exists(last_out['path'])
                       else messagebox.showinfo('提示', '还没有结果文件'))
    btn_rev.configure(command=lambda: reveal_path(last_out['path'])
                      if last_out['path'] and os.path.exists(last_out['path'])
                      else messagebox.showinfo('提示', '还没有结果文件'))

    if preset and os.path.isdir(preset):
        win.after(200, on_change)

    hint_bar(host, '「写入标黄」会直接修改原表；动手前建议先用「预览」空跑确认。', F).pack(
        anchor='w', fill='x')
    fade_in(win)
    return win
