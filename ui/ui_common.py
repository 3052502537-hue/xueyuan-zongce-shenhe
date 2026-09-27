# -*- coding: utf-8 -*-
"""
ui_common · 统一门面模块（兼容层）
============================================================================
本模块只做两件事：
    1. setup_style(root) —— ttk 全局主题（必须在其他 UI 模块加载前调用）
    2. 把所有 ui_design / ui_render / ui_window / ui_widgets / ui_layouts /
       ui_animations / ui_runner 的公共接口 re-export 出来，业务代码只需
       `from ui.ui_common import ...` 即可，不需要知道下面细分了 7 个子模块。

改外观 / 改交互细节时去对应的子模块，单点修改即可。

模块清单（均在 ui/ 包内）：
    ui_design        颜色、字体、几何常量
    ui_render        Pillow 渲染（圆角图片、阴影、logo）
    ui_window        平台工具、窗口拖动、置顶、放置
    ui_widgets       圆角按钮、卡片、勾选框、输入框、状态胶囊、表格
    ui_layouts       顶部导航、滚动容器、路径行、提示条
    ui_animations    淡入、进度条过渡、脉冲、抖动、数字滚动
    ui_runner        Runner 后台任务驱动、QueueWriter、日志区
"""
import tkinter as tk
from tkinter import ttk

from .ui_design import (
    # 平台判断
    IS_MAC, IS_WIN,
    # 颜色
    C_ACCENT, C_ACCENT_HOVER, C_ACCENT_SOFT, C_ACCENT_TINT,
    C_PAGE, C_SURFACE, C_SUBTLE, C_CARD_HOVER,
    C_BORDER, C_BORDER_STRONG, C_BORDER_HOVER, C_TREE_HEAD,
    C_TEXT, C_MUTED, C_HINT,
    C_OK, C_WARN, C_ERR,
    C_DISABLED_BG, C_DISABLED_FG, C_SHADOW,
    # 几何
    RADIUS_CARD, RADIUS_BTN, RADIUS_TAG, RADIUS_INPUT,
    SHADOW_M, SHADOW_BLUR, SHADOW_OFF, SHADOW_ALPHA,
    # 字体
    Fonts,
)
from .ui_render import (
    HAVE_PIL,
    card_photo, btn_photo, pill_photo, logo_photo, clear_cache,
)
from .ui_window import (
    SCALE, sp,
    downloads_dir, default_out_root, open_path, reveal_path,
    load_module, load_engine,
    make_borderless, enable_drag,
    WIN_MIN, WIN_MAX, WIN_CLOSE,
    place_window, force_foreground, topmost_once,
)
from .ui_layouts import (
    header_bar, hint_bar, section_label, page_label, muted_label,
    ScrollableFrame, path_row, page_scaffold, action_bar,
)
from .ui_widgets import (
    AnimButton,
    primary_button, secondary_button, ghost_button,
    StatusPill,
    RoundedCard,
    card_section, stat_card,
    enable_tree_hover, scroll_table,
    CheckBox, checkbox,
    RoundedEntry, entry,
)
from .ui_animations import (
    fade_in, animate_progress, stagger_highlight,
    pulse, shake, count_up,
)
from .ui_runner import (
    Runner, QueueWriter, LogPane,
)


# ============================================================
# ttk 全局主题（这是 ui_common.py 唯一保留的"自有"职责）
# ============================================================
def setup_style(root):
    """声明 DPI + 安装缩放补丁 + 配置 ttk 主题。必须在 root 创建后调用一次。"""
    from .ui_window import _init_scale, _install_scaling_patches
    _init_scale(root)
    _install_scaling_patches()
    try:
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('TFrame', background=C_SURFACE)
        s.configure('TLabel', background=C_SURFACE, foreground=C_TEXT)
        s.configure('TLabelFrame', background=C_SURFACE, bordercolor=C_BORDER,
                    relief='solid', borderwidth=1,
                    labelmargins=(10, 4))
        s.configure('TLabelFrame.Label', background=C_SURFACE,
                    foreground=C_TEXT)
        s.configure('TEntry', fieldbackground=C_SURFACE, foreground=C_TEXT,
                    bordercolor=C_BORDER_STRONG, lightcolor=C_BORDER_STRONG,
                    darkcolor=C_BORDER_STRONG, insertcolor=C_ACCENT,
                    padding=6)
        s.configure('TCheckbutton', background=C_SURFACE, foreground=C_TEXT,
                    focuscolor=C_SURFACE, indicatorcolor=C_SURFACE,
                    bordercolor=C_BORDER_STRONG, padding=(2, 4))
        s.map('TCheckbutton',
              background=[('active', C_SURFACE)],
              indicatorcolor=[('selected', C_ACCENT), ('!selected', C_SURFACE)],
              foreground=[('disabled', C_HINT)])
        s.configure('TButton', background=C_SURFACE, foreground=C_TEXT,
                    bordercolor=C_BORDER_STRONG, relief='flat', focusthickness=0,
                    padding=(10, 6))
        s.map('TButton', background=[('active', C_SUBTLE),
                                     ('pressed', C_SUBTLE)])
        s.configure('TProgressbar', troughcolor='#F0F0F0', background=C_ACCENT,
                    bordercolor='#F0F0F0', lightcolor=C_ACCENT,
                    darkcolor=C_ACCENT, borderwidth=0, thickness=8)
        s.configure('Treeview', background=C_SURFACE,
                    fieldbackground=C_SURFACE, foreground=C_TEXT,
                    rowheight=34, borderwidth=0, relief='flat')
        s.map('Treeview', background=[('selected', C_ACCENT_SOFT)],
              foreground=[('selected', C_ACCENT)])
        s.configure('Treeview.Heading', background=C_TREE_HEAD,
                    foreground=C_MUTED, relief='flat', borderwidth=0,
                    padding=(8, 9))
        s.map('Treeview.Heading', background=[('active', '#EDEFF2')])
        s.configure('TNotebook', background=C_SURFACE, borderwidth=0,
                    tabmargins=(0, 4, 0, 0))
        s.configure('TNotebook.Tab', background=C_SUBTLE, foreground=C_MUTED,
                    padding=(16, 7), borderwidth=0)
        s.map('TNotebook.Tab', background=[('selected', C_SURFACE)],
              foreground=[('selected', C_ACCENT)])
        s.configure('Vertical.TScrollbar', background='#DCDEE3',
                    troughcolor=C_SURFACE, bordercolor=C_SURFACE,
                    arrowcolor=C_MUTED, relief='flat', borderwidth=0,
                    arrowsize=14)
        s.configure('Horizontal.TScrollbar', background='#DCDEE3',
                    troughcolor=C_SURFACE, bordercolor=C_SURFACE,
                    arrowcolor=C_MUTED, relief='flat', borderwidth=0,
                    arrowsize=14)
    except Exception:
        pass
    try:
        root.configure(bg=C_PAGE)
    except Exception:
        pass


# ============================================================
# __all__
# ============================================================
__all__ = [
    # 平台 / 缩放
    'IS_MAC', 'IS_WIN', 'SCALE', 'sp',
    # 颜色
    'C_ACCENT', 'C_ACCENT_HOVER', 'C_ACCENT_SOFT', 'C_ACCENT_TINT',
    'C_PAGE', 'C_SURFACE', 'C_SUBTLE', 'C_CARD_HOVER',
    'C_BORDER', 'C_BORDER_STRONG', 'C_BORDER_HOVER', 'C_TREE_HEAD',
    'C_TEXT', 'C_MUTED', 'C_HINT',
    'C_OK', 'C_WARN', 'C_ERR',
    'C_DISABLED_BG', 'C_DISABLED_FG', 'C_SHADOW',
    # 几何
    'RADIUS_CARD', 'RADIUS_BTN', 'RADIUS_TAG', 'RADIUS_INPUT',
    'SHADOW_M', 'SHADOW_BLUR', 'SHADOW_OFF', 'SHADOW_ALPHA',
    # 字体
    'Fonts',
    # 平台工具
    'downloads_dir', 'default_out_root', 'open_path', 'reveal_path',
    'load_module', 'load_engine',
    # 窗口
    'make_borderless', 'enable_drag', 'place_window', 'topmost_once',
    'force_foreground',
    # 渲染
    'HAVE_PIL', 'card_photo', 'btn_photo', 'pill_photo', 'logo_photo',
    'clear_cache',
    # 控件
    'AnimButton', 'primary_button', 'secondary_button', 'ghost_button',
    'StatusPill',
    'RoundedCard', 'card_section', 'stat_card',
    'CheckBox', 'checkbox',
    'RoundedEntry', 'entry',
    'scroll_table', 'enable_tree_hover',
    # 布局
    'header_bar', 'hint_bar', 'section_label', 'page_label', 'muted_label',
    'ScrollableFrame', 'path_row', 'page_scaffold', 'action_bar',
    # 动效
    'fade_in', 'animate_progress', 'stagger_highlight',
    'pulse', 'shake', 'count_up',
    # 后台 / 日志
    'Runner', 'QueueWriter', 'LogPane',
    # 主题（门面模块自有）
    'setup_style',
]