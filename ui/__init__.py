# -*- coding: utf-8 -*-
"""
ui · 综测核算程序界面层（独立包）
============================================================================
把全部界面相关模块收拢到 ui/ 包内，与业务引擎（verify.py / 挂科引擎.py /
回填引擎.py）和任务页（task_*.py）在目录上彻底分开，便于后续单独维护、
升级或替换整套 UI。

对外统一入口仍是 ``ui.ui_common``（兼容门面），业务代码只需：

    from ui.ui_common import (Fonts, LogPane, Runner, ...)

包内模块一律使用相对导入（``from .ui_design import ...``），彼此不再依赖
项目根目录在 sys.path 上。

子模块职责：
    ui_design       颜色、字体、几何常量
    ui_render       Pillow 渲染（圆角图片、阴影、logo）
    ui_window       平台工具、窗口拖动、置顶、放置、引擎加载
    ui_widgets      圆角按钮、卡片、勾选框、输入框、状态胶囊、表格
    ui_layouts      顶部导航、滚动容器、路径行、提示条
    ui_animations   淡入、进度条过渡、脉冲、抖动、数字滚动
    ui_runner       Runner 后台任务驱动、QueueWriter、日志区
    ui_common       门面：re-export 上述公共接口 + setup_style
"""
