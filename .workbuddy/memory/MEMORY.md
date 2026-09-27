# 项目长期记忆 · 综测核算程序

## 项目性质
太原科技大学计算机科学与技术学院的成绩处理工具（Python + tkinter + openpyxl）。
三任务：① 综测两表核对（verify.py）② 挂科/附加分筛查（挂科引擎.py）③ 回填综测总表（回填引擎.py）。
下游契约：K 列是否挂科来自挂科汇总；E 列附加分必须填 0 而非留空（F=D+E 依赖）。

## 代码结构约定
- 入口 `main.py` + 三个任务页 `task_*.py` 只做调度，规则全在引擎（项目根）。
- 所有 UI 收拢在 `ui/` 包：对外统一 `from ui.ui_common import ...`；包内互相用相对导入。
- `ui/ui_common.py` 是门面（re-export + setup_style），改外观去对应子模块单点修改。
- 引擎经 `ui/ui_window.py::load_engine` 按 `here`（任务页所在项目根）路径加载，留在项目根不动。

## 打包 / 签名约定
- 一键：`python build.py`（PyInstaller 打包 → `sign_exe.py` 签名）。
- 版本资源 `version_info.txt` 必须是返回 `VSVersionInfo(...)` 的**纯表达式**（PyInstaller `eval`，禁写 import/赋值）。
- 签名：`sign_exe.py` 默认生成 CN=程震宇 自签名证书（sigtool 走 Windows SDK x64）；
  正规 CA 证书用 `SIGN_PFX`/`SIGN_PWD`。安装包见 `installer.iss`（需 Inno Setup 的 ISCC）。
- 运行环境：用**商店版 Python 3.11.9**（自带 tkinter/openpyxl/pillow）；托管 Python 3.13.12 无 tkinter。

## 验收基线（移植/改动后必须能对上）
② 全量：挂科 550 人 / 1042 门次；附加分 29 人 / 30 门次。
③ 各年级总表：23 级 592 行(62是/530否/9人附加分)、24 级 667(181/486/8)、25 级 751(302/449/12人13门次)。
