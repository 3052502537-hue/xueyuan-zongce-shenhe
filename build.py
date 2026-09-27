# -*- coding: utf-8 -*-
"""
build · 一键构建（PyInstaller 打包 → 数字签名）
============================================================================
等价于：
    pyinstaller 综测核算程序.spec        # 先打包（已内嵌版本资源：作者程震宇 / v1.0）
    python sign_exe.py dist/综测核算程序.exe   # 再签名（签名者 = 程震宇）

用法：
    python build.py
    SIGN_PFX=my.pfx SIGN_PWD=123456 python build.py   # 用正规证书签名

产物：dist/综测核算程序.exe（已签名）。安装包另见 installer.iss（需 Inno Setup）。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def build():
    from PyInstaller.__main__ import run
    spec = os.path.join(HERE, '综测核算程序.spec')
    print('[构建] PyInstaller 打包：%s' % spec)
    run(['--noconfirm', spec])


def sign():
    exe = os.path.join(HERE, 'dist', '综测核算程序.exe')
    if not os.path.isfile(exe):
        print('[签名] 未找到构建产物：%s，跳过签名' % exe, file=sys.stderr)
        return False
    import sign_exe
    return sign_exe.sign_file(exe)


def main():
    build()
    ok = sign()
    print('\n完成。产物：%s' % os.path.join(HERE, 'dist', '综测核算程序.exe'))
    if not ok:
        print('（警告：签名未成功，见上方日志；可手动运行 python sign_exe.py 重试）')
        sys.exit(1)


if __name__ == '__main__':
    main()
