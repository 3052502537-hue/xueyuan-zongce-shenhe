# -*- coding: utf-8 -*-
"""
ui_design · 设计令牌与字体
============================================================================
整个 UI 系统的「设计契约」：所有颜色、几何、阴影参数、字体规格都在这里定义。
其他 UI 模块从这里导入；改主题只动这里。

配色：主色 #1890FF ｜ 辅助浅灰 #F5F7FA ｜ 正文 #333333 ｜ 次要 #888888
纯白背景 + 浅灰细边框 + 圆角；字体 Microsoft YaHei UI / Segoe UI。
"""
import platform
import tkinter as tk
import tkinter.font as tkfont


IS_MAC = platform.system() == 'Darwin'
IS_WIN = platform.system() == 'Windows'


# ============================================================
# 颜色令牌（设计系统级常量；按用途分组）
# ============================================================
# —— 主色 / 强调 ——
C_ACCENT        = '#1890FF'
C_ACCENT_HOVER  = '#096DD9'
C_ACCENT_SOFT   = '#E6F4FF'
C_ACCENT_TINT   = '#F0F7FF'

# —— 中性面 ——
C_PAGE          = '#FFFFFF'
C_SURFACE       = '#FFFFFF'
C_SUBTLE        = '#F5F7FA'
C_CARD_HOVER    = '#F5F7FA'

# —— 边框 ——
C_BORDER        = '#EBEEF5'
C_BORDER_STRONG = '#D9D9D9'
C_BORDER_HOVER  = '#1890FF'
C_TREE_HEAD     = '#F5F7FA'

# —— 文字 ——
C_TEXT          = '#333333'
C_MUTED         = '#888888'
C_HINT          = '#B0B0B0'

# —— 状态 ——
C_OK            = '#52C41A'
C_WARN          = '#FAAD14'
C_ERR           = '#FF4D4F'

# —— 其它 ——
C_DISABLED_BG   = '#F5F5F5'
C_DISABLED_FG   = '#BFBFBF'
C_SHADOW        = '#EDF0F5'

# —— 进度 / 胶囊背景（饱和度更低）——
BG_OK_SOFT      = '#F6FFED'
BG_ERR_SOFT     = '#FFF1F0'


# ============================================================
# 几何令牌
# ============================================================
RADIUS_CARD   = 12    # 卡片圆角
RADIUS_BTN    = 10    # 按钮圆角
RADIUS_TAG    = 6     # 标签圆角
RADIUS_INPUT  = 8     # 输入框圆角

# 阴影参数（传给 Pillow 渲染）
SHADOW_M      = 7     # 阴影边距
SHADOW_BLUR   = 5     # 模糊半径
SHADOW_OFF    = 3     # 偏移
SHADOW_ALPHA  = 24    # 透明度


# ============================================================
# 字体
# ============================================================
def _ui_family():
    """返回当前平台推荐的无衬线字体族。"""
    avail = set(tkfont.families())
    if IS_WIN:
        for f in ('Microsoft YaHei UI', 'Microsoft YaHei', 'Segoe UI'):
            if f in avail:
                return f
    elif IS_MAC:
        for f in ('PingFang SC', 'Heiti SC'):
            if f in avail:
                return f
    return 'TkDefaultFont'


def _ui_mono():
    """返回当前平台推荐的等宽字体族。"""
    avail = set(tkfont.families())
    if IS_WIN:
        return 'Consolas' if 'Consolas' in avail else 'TkFixedFont'
    if IS_MAC:
        return 'Menlo' if 'Menlo' in avail else 'TkFixedFont'
    return 'TkFixedFont'


class Fonts(object):
    """统一字体规格：UI 字体族 + 等宽字体族 + 各层级字号的 Font 实例。

    使用方式：
        F = Fonts()                       # 主入口创建一次
        tk.Label(..., font=F.h2)          # 直接拿某个层级的 Font
        tk.Label(..., font=F.small)       # 字号详见 self._make()
    """
    _INSTANCES = {}     # 进程内单例：避免重复创建 tkfont.Font 实例

    def __new__(cls):
        if cls not in cls._INSTANCES:
            inst = super().__new__(cls)
            inst._init()
            cls._INSTANCES[cls] = inst
        return cls._INSTANCES[cls]

    def _init(self):
        self.fam = _ui_family()
        self.mono = _ui_mono()
        self.h1      = tkfont.Font(family=self.fam, size=17, weight='bold')
        self.title   = tkfont.Font(family=self.fam, size=16, weight='bold')
        self.h2      = tkfont.Font(family=self.fam, size=12, weight='bold')
        self.sec     = tkfont.Font(family=self.fam, size=11, weight='bold')
        self.normal  = tkfont.Font(family=self.fam, size=11)
        self.small   = tkfont.Font(family=self.fam, size=10)
        self.monos   = tkfont.Font(family=self.mono, size=10)
        self.card_t  = tkfont.Font(family=self.fam, size=14, weight='bold')
        self.stat    = tkfont.Font(family=self.fam, size=34, weight='bold')

    @classmethod
    def reset(cls):
        """重置单例（仅供测试）。"""
        cls._INSTANCES.clear()