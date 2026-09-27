# -*- mode: python ; coding: utf-8 -*-
import os

HERE = SPECPATH
# 三个引擎经 load_engine() 按文件路径加载，必须作为数据文件内嵌到 exe 中
# 界面模块已整体收入 ui/ 包；PyInstaller 从 main.py 的 `from ui.ui_common import`
# 自动追踪并收集整个 ui 包，无需在此显式列出。

ENGINE_FILES = ['verify.py', '挂科引擎.py', '回填引擎.py']
datas = []
for _f in ENGINE_FILES:
    _p = os.path.join(HERE, _f)
    if os.path.isfile(_p):
        datas.append((_p, '.'))


a = Analysis(
    [os.path.join(HERE, 'main.py')],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=['openpyxl', 'PIL', 'PIL.Image', 'PIL.ImageDraw',
                   'PIL.ImageFilter', 'PIL.ImageTk', 'PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'lxml'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='综测核算程序',
    # 版本资源：作者=程震宇、版本=1.0（文件属性里可见）
    version=os.path.join(HERE, 'version_info.txt'),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
