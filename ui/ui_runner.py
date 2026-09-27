# -*- coding: utf-8 -*-
"""
ui_runner · 后台任务驱动
============================================================================
把"子线程工作 + 主线程更新 UI"这件容易写错的事情封装好：

    runner = Runner(root, logpane, handler=, on_finish=)
    runner.run(lambda q: worker(q, snapshot))

子线程 worker 通过 q.put((kind, payload)) 投递消息：
    ('log', '...')             → 写入日志区
    ('done', None)              → 任务结束
    其它 kind                  → 转给 handler(kind, payload)（界面自定义处理）
完成时调用 on_finish()。

公开 API：
    Runner              主类
    QueueWriter         可配合 contextlib.redirect_stdout 把 print 输出塞进队列

⚠ 使用要点：
    1) 子线程不能碰任何 tk 变量 / 控件——会抛 "main thread is not in main loop"
    2) 所有界面取值都要在主线程里 snapshot 成普通 dict 后传给 worker
    3) 一个 Runner 同时只跑一个任务（busy 状态保护）
"""
import queue
import threading
import traceback

import tkinter as tk
from tkinter import ttk, scrolledtext

from .ui_design import (
    C_PAGE, C_TEXT, C_ACCENT, C_BORDER,
    C_ERR, C_OK, C_WARN, C_HINT,
)
from .ui_window import IS_WIN


# ============================================================
# Runner：子线程 + 主线程轮询
# ============================================================
class Runner(object):
    """统一的「后台任务 → 主线程更新」驱动器。

    用法：
        runner = Runner(root, logpane, handler=, on_finish=)
        def worker(q, cfg): ...   # 子线程：q.put(('log', ...)) 等
        runner.run(lambda q: worker(q, cfg))

    handler(kind, payload)：主线程收到「非 log / 非 done」消息时被调用。
        用于更新 stat_card / treeview 等 Tk 控件。
    on_finish()：任务结束（无论成败）后被调用一次。
    """

    def __init__(self, root, logpane, handler=None, on_finish=None):
        self.root = root
        self.log = logpane
        self.handler = handler
        self.on_finish = on_finish
        self.busy = False
        self.q = queue.Queue()

    def run(self, worker):
        """提交一个 worker(q) 函数到后台线程。

        返回 True 表示成功启动；False 表示当前已有任务在跑。
        """
        if self.busy:
            return False
        self.busy = True
        threading.Thread(target=self._wrap, args=(worker,), daemon=True).start()
        self._poll()
        return True

    def _wrap(self, worker):
        """子线程入口：捕获异常并投递 'done' 哨兵。"""
        try:
            worker(self.q)
        except Exception as e:
            self.q.put(('log', '✘ 出错：%s' % e))
            self.q.put(('log', traceback.format_exc(limit=4)))
        finally:
            self.q.put(('done', None))

    def _poll(self):
        """主线程轮询：把所有 pending 消息分发给 log / handler。"""
        done = False
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'done':
                    done = True
                elif kind == 'log':
                    self.log.write(payload)
                elif self.handler is not None:
                    self.handler(kind, payload)
        except queue.Empty:
            pass
        if done:
            self.busy = False
            if self.on_finish is not None:
                self.on_finish()
            return
        self.root.after(100, self._poll)


# ============================================================
# QueueWriter：把 print() 的输出塞进队列
# ============================================================
class QueueWriter(object):
    """文件式接口（.write/.flush），把 print 输出转成 ('log', line) 投进队列。

    用法（常见模式）：
        with contextlib.redirect_stdout(QueueWriter(q)):
            eng.verify_class(...)
    """

    def __init__(self, q):
        self.q = q

    def write(self, s):
        for line in str(s).splitlines():
            if line.strip():
                self.q.put(('log', line.rstrip()))
        return len(s)

    def flush(self):
        pass


# ============================================================
# 日志区（带颜色 tag 的滚动文本）
# ============================================================
class LogPane(ttk.Frame):
    """带颜色 tag 的滚动文本日志区。

    tag：
        err   错误（红色）
        ok    成功（绿色）
        warn  警告（橙色）
        dim   提示（浅灰）
    """

    def __init__(self, parent, height=12):
        ttk.Frame.__init__(self, parent)
        mono = 'Consolas' if IS_WIN else 'Menlo'
        self.txt = scrolledtext.ScrolledText(
            self, height=height, wrap='none',
            font=(mono, 10), bg='#FAFAFA', fg=C_TEXT,
            insertbackground=C_ACCENT, relief='flat', borderwidth=0,
            padx=10, pady=8, highlightthickness=1,
            highlightbackground=C_BORDER, highlightcolor=C_ACCENT)
        self.txt.pack(fill='both', expand=True)
        self.txt.tag_config('err', foreground=C_ERR)
        self.txt.tag_config('ok', foreground=C_OK)
        self.txt.tag_config('dim', foreground=C_HINT)
        self.txt.tag_config('warn', foreground=C_WARN)

    def write(self, msg, tag=None):
        if tag is None:
            low = str(msg)
            if '✘' in low or '错误' in low or '失败' in low or 'Traceback' in low:
                tag = 'err'
            elif low.startswith('✔') or low.startswith('✓'):
                tag = 'ok'
        self.txt.insert('end', msg + '\n', tag)
        self.txt.see('end')

    def clear(self):
        self.txt.delete('1.0', 'end')