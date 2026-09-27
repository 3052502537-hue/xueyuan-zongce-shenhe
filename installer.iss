; 综测核算程序 · Inno Setup 安装包脚本
; ============================================================================
; 编译前请先安装 Inno Setup（https://jrsoftware.org/isdl.php），然后在 Inno
; Setup Compiler 中打开本文件点「Compile」，或命令行：
;     "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
;
; 前提：先运行 `python build.py` 生成 dist\综测核算程序.exe（已签名）。
; 安装包输出到项目根：综测核算程序_Setup.exe
; ============================================================================

[Setup]
; --- 基本信息 ---
AppName=综测核算程序
AppVersion=1.0
AppPublisher=程震宇
AppPublisherURL=https://www.tyust.edu.cn
AppId={{A1B2C3D4-2026-4E01-8C01-000000000001}
DefaultDirName={autopf}\综测核算程序
DefaultGroupName=综测核算程序
OutputBaseFilename=综测核算程序_Setup
OutputDir=.
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64os
; 对卸载程序也签名（需要 signtool.exe 在 PATH；若没有，改 no 或把 SDK 目录加入 PATH）
SignedUninstaller=no
; 注意：安装包本身不含驱动/服务，默认以普通用户权限安装即可
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\综测核算程序.exe

[Languages]
Name: "chinese"; MessagesFile: "compiler:Default.isl"

[Files]
; 主体：已签名的单文件 exe
Source: "dist\综测核算程序.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\综测核算程序"; Filename: "{app}\综测核算程序.exe"; \
  WorkingDir: "{app}"
Name: "{group}\卸载 综测核算程序"; Filename: "{uninstallexe}"
; 桌面快捷方式（默认勾选，可在安装向导里取消）
Name: "{autodesktop}\综测核算程序"; Filename: "{app}\综测核算程序.exe"; \
  Tasks: desktopicon; WorkingDir: "{app}"

[Tasks]
Name: desktopicon; Description: "创建桌面快捷方式(&D)"; \
  GroupDescription: "附加选项："; Flags: unchecked

[Run]
; 安装结束可选启动
Filename: "{app}\综测核算程序.exe"; Description: "启动 综测核算程序"; \
  Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 清理可能的运行残留（不影响用户数据）
Type: files; Name: "{app}\selftest_error.txt"
