# -*- coding: utf-8 -*-
"""
sign_exe · Windows Authenticode 数字签名工具
============================================================================
对打包好的 exe 进行代码签名，签名者显示为「程震宇」。两种模式：

  1) 正规代码签名证书（推荐；Windows 显示为可信发布者"程震宇"，无 SmartScreen 拦截）：
       设置环境变量后运行：
         set SIGN_PFX=你的证书.pfx
         set SIGN_PWD=证书密码
         python sign_exe.py dist/综测核算程序.exe
     还可选设置 SIGN_TIMESTAMP=RFC3161 时间戳地址（让签名在证书过期后仍有效）。

  2) 自签名证书（默认；签名真实有效，但 Windows 会提示"发布者未验证"）：
       不设置 SIGN_PFX 时，自动用 PowerShell 生成 CN=程震宇 的代码签名证书，
       导出到项目根 dev_cert.pfx 后用 signtool 签名。适合内部/测试分发。
       想换成正规证书时，只改环境变量即可，无需改本脚本。

依赖：
   - signtool.exe（Windows SDK 自带，本脚本会在常见路径与 PATH 中自动查找）
   - 自签名模式下还需 PowerShell（New-SelfSignedCertificate / Export-PfxCertificate）

典型调用：
   python sign_exe.py                # 默认对 dist/综测核算程序.exe 自签名
   python sign_exe.py <exe路径>      # 指定文件
   python sign_exe.py <exe路径> --timestamp http://timestamp.digicert.com
"""
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))

# 默认自签名证书存放位置（仅自签名模式生成，可随时删除）
DEFAULT_SELF_PFX = os.path.join(HERE, 'dev_cert.pfx')
DEFAULT_SELF_PWD = 'devpass'


def find_signtool():
    """在 PATH 与 Windows SDK 常见路径里定位 signtool.exe。

    优先 x64 / x86 版本，排除 arm / arm64（后者在 x64 Windows 上会报
    WinError 216 "版本不兼容"）。
    """
    env = os.environ.get('SIGNTOOL')
    if env and os.path.isfile(env):
        return env
    sdk_root = r'C:\Program Files (x86)\Windows Kits\10\bin'
    found = []
    if os.path.isdir(sdk_root):
        for root, _dirs, files in os.walk(sdk_root):
            if 'signtool.exe' in files:
                found.append(os.path.join(root, 'signtool.exe'))
    # 优先 x64，其次 x86
    for arch in ('x64', 'x86'):
        for p in found:
            if ('\\%s\\' % arch) in p.replace('/', '\\'):
                return p
    # 兜底：排除 arm / arm64
    for p in found:
        low = p.lower().replace('/', '\\')
        if '\\arm' not in low:
            return p
    # 退回 PATH
    from shutil import which
    p = which('signtool.exe') or which('signtool')
    if p:
        return p
    return None


def _run_powershell(script):
 """运行一段 PowerShell 脚本，返回 (returncode, 文本输出)。

 显式把输出编码设为 UTF-8，否则中文（如签名者"程震宇"）会以系统代码页
 (GBK) 输出，导致 Python 按 UTF-8 解码失败。
 """
 enc = ('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; '
 '$OutputEncoding = [System.Text.Encoding]::UTF8; ')
 script = enc + script
 proc = subprocess.run(
     ['powershell', '-NoProfile', '-NonInteractive', '-Command', script],
     capture_output=True,
 )
 raw = (proc.stdout or b'') + (proc.stderr or b'')
 try:
     out = raw.decode('utf-8')
 except Exception:
     out = raw.decode('utf-8', 'replace')
 return proc.returncode, out


def make_self_signed_pfx(subject='程震宇', pfx_path=DEFAULT_SELF_PFX,
                        password=DEFAULT_SELF_PWD):
    """用 PowerShell 生成代码签名证书并导出为 pfx；返回 pfx 路径。"""
    script = (
        "$ErrorActionPreference='Stop';"
        # 若已存在同名 pfx 且有效则直接复用
        "if (Test-Path '%s') { Write-Host 'REUSE_PFX'; exit 0 }"
        "$c = New-SelfSignedCertificate -Type CodeSigningCert "
        "-Subject 'CN=%s' -CertStoreLocation 'Cert:\\CurrentUser\\My' "
        "-KeySpec KeyExchange;"
        "$pw = ConvertTo-SecureString -String '%s' -Force -AsPlainText;"
        "Export-PfxCertificate -Cert $c -FilePath '%s' -Password $pw | Out-Null;"
        # 注意：保留证书在个人证书库（Cert:\CurrentUser\My），便于本机验证签名；"
        # 仅 pfx 文件留在项目根（dev_cert.pfx）可随时删除重建。"
        "Write-Host 'MADE_PFX';"
    ) % (pfx_path.replace('/', '\\'),
         subject,
         password,
         pfx_path.replace('/', '\\'))
    rc, out = _run_powershell(script)
    if rc != 0:
        raise RuntimeError('生成自签名证书失败：\n' + out)
    if not os.path.isfile(pfx_path):
        raise RuntimeError('pfx 未生成：\n' + out)
    return pfx_path


def sign_file(exe_path, timestamp=None):
    """对 exe 签名：优先用 SIGN_PFX/SIGN_PWD，否则自签名。返回是否成功。"""
    signtool = find_signtool()
    if not signtool:
        print('[签名] 未找到 signtool.exe（需 Windows SDK）。签名跳过。', file=sys.stderr)
        print('        可安装 Windows SDK 或将 signtool 所在目录加入 PATH 后重试。', file=sys.stderr)
        return False

    pfx = os.environ.get('SIGN_PFX')
    pwd = os.environ.get('SIGN_PWD')
    if pfx and os.path.isfile(pfx):
        print('[签名] 使用正规证书：%s' % pfx)
    else:
        print('[签名] 未提供 SIGN_PFX，改用自签名证书（CN=程震宇）。')
        pfx = make_self_signed_pfx()
        pwd = DEFAULT_SELF_PWD

    cmd = [signtool, 'sign', '/fd', 'SHA256']
    if timestamp:
        cmd += ['/tr', timestamp, '/td', 'SHA256']
    cmd += ['/f', pfx, '/p', pwd, exe_path]
    print('[签名] 执行：%s' % ' '.join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True)
    print(proc.stdout)
    if proc.stderr:
        print(proc.stderr)
    if proc.returncode != 0:
        print('[签名] 失败（退出码 %d）' % proc.returncode, file=sys.stderr)
        return False

    # 验证（用 PowerShell 的 Authenticode 检查，能直接读出签名者名称）。
    # 注意：自签名证书会因"根证书不受信任"而 Status=UnknownError，但签名本身
    # 完整有效（HASSIG=True）。这里以"签名存在 + 签名者=程震宇"判定成功，
    # 信任与否另行说明；signtool verify /pa 也会因缺 CA 链而误报，故不采用。
    script = (
        "$s = Get-AuthenticodeSignature -FilePath '%s';"
        "$ok = ($s.SignerCertificate -ne $null);"
        "Write-Output ('STATUS=' + $s.Status);"
        "Write-Output ('HASSIG=' + $ok.ToString());"
        "Write-Output ('SUBJECT=' + $s.SignerCertificate.Subject);"
    ) % exe_path.replace('/', '\\')
    rc, out = _run_powershell(script)
    print('[签名] 验证输出：\n' + out)
    import re
    m = re.search(r'STATUS=(\S+)', out)
    status = m.group(1) if m else ''
    m = re.search(r'HASSIG=(\w+)', out)
    has_sig = (m and m.group(1) == 'True')
    m = re.search(r'SUBJECT=(.+)', out)
    subject = m.group(1).strip() if m else ''
    ok = bool(has_sig) and ('程震宇' in subject)
    if status == 'Valid':
        print('[签名] 完成 ✅（签名有效且受系统信任）')
    elif ok:
        print('[签名] 完成 ✅（已签名，签名者 = %s；自签名未经 CA 信任属正常，换正规证书即受信任）' % subject)
    else:
        print('[签名] 验证未通过 ❌（status=%s）' % status)
    return ok


def main():
    args = sys.argv[1:]
    exe_path = None
    timestamp = os.environ.get('SIGN_TIMESTAMP')
    if args and not args[0].startswith('--'):
        exe_path = args.pop(0)
    for i, a in enumerate(args):
        if a == '--timestamp' and i + 1 < len(args):
            timestamp = args[i + 1]
    if not exe_path:
        exe_path = os.path.join(HERE, 'dist', '综测核算程序.exe')
    if not os.path.isfile(exe_path):
        print('未找到待签名文件：%s' % exe_path, file=sys.stderr)
        sys.exit(2)
    ok = sign_file(exe_path, timestamp=timestamp)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
