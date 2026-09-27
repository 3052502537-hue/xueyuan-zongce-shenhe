# Windows 移植说明（macOS 版 → Windows 版）

> 这份文档写给**拿到本程序、要在 Windows 上做适配的 AI（或同学）**。
> 目标：不改任何业务规则，只做平台适配，就能在 Windows 上跑出与 macOS **完全一致**的结果。

---

## 0. 一句话结论

**业务代码（4 个 `.py` 引擎 + 界面逻辑）是纯 Python + openpyxl + tkinter，本身跨平台，不需要改。**
真正需要适配的只有 3 处：

| # | 位置 | macOS 现状 | Windows 要改成 |
|---|---|---|---|
| 1 | `启动.command` | bash 脚本，双击/拖拽启动 | 换成 `启动.bat`（或 `.ps1`） |
| 2 | `ui/ui_window.py` | `downloads_dir()` / `open_path()` / `reveal_path()` 用 `open` 命令 | 用 `os.startfile` / `explorer /select,` |
| 3 | `ui/ui_design.py` → `Fonts` | 字体 `PingFang SC` / `Menlo` | 字体 `Microsoft YaHei UI` / `Consolas` |

其余（判定规则、写入逻辑、学号处理、报告格式）**照抄即可**，下面第 6 节会证明为什么不用改。

---

## 1. 这个程序是干什么的

太原科技大学计算机科学与技术学院的成绩处理工具，一个入口三个任务：

| 任务 | 作用 | 输入 | 产出 |
|---|---|---|---|
| **① 综测两表核对** | 校验「综合素质测评详情表」与「综合素质测评表」两份表是否一致 | 年级文件夹（每班一个子文件夹，含上述两个 xlsx） | 核验报告 xlsx（**只读源表，绝不写回**） |
| **② 挂科/附加分筛查** | 识别必修课 → 判定挂科学生（可整行标黄）→ 认定必修课 100 分附加分资格 | 年级文件夹 或 学期根目录（内含各班 `班级成绩(xx).xlsx`） | 《挂科与附加分表.xlsx》（7 个 Sheet） |
| **③ 回填成绩表** | 依据 ② 的产出，回填各年级「综测成绩表」的**是否挂科**与**附加分** | 综测总表目录 + ②产出的 xlsx | 回填核验报告 md（**并直接写入总表**） |

**③ 依赖 ②**：③ 的数据源就是②产出的《挂科与附加分表.xlsx》。这个联系不能断。

---

## 2. 目录结构

```
main.py              首页：三张任务卡 + 产物输出根目录
ui/                  【界面层独立包】所有 UI 模块收拢在此
  ├─ ui_common.py    门面（对外入口）：re-export 公共接口 + setup_style
  ├─ ui_design.py    设计令牌：Fonts / 颜色 / 几何常量
  ├─ ui_window.py    窗口工具：平台路径 / 拖动 / 置顶 / load_engine
  ├─ ui_widgets.py   自定义控件：按钮 / 卡片 / 勾选框 / 输入框 / 表格
  ├─ ui_layouts.py   布局组件：头部导航 / 滚动容器 / 路径行
  ├─ ui_render.py    Pillow 渲染：圆角图片 / 阴影 / logo
  ├─ ui_animations.py 动效：淡入 / 数字滚动 / 进度 / 脉冲 / 抖动
  └─ ui_runner.py    后台任务驱动：Runner / LogPane / QueueWriter
task_verify.py       ① 界面（构建函数 build）
task_screening.py    ② 界面
task_fillback.py     ③ 界面
verify.py            ① 引擎（两表核对规则，md5 固定，规则一字未改）
挂科引擎.py          ② 引擎（= 生成挂科附加分表.py 的副本）
回填引擎.py          ③ 引擎（CLI 也能跑）
启动.command         macOS 启动器（双击 / 拖文件夹到图标上）
```

> 注意：UI 模块现已整体收入 `ui/` 包（包内用相对导入 `from .ui_xxx import`），
> 业务代码通过 `from ui.ui_common import ...` 取用。引擎文件仍留在项目根目录，
> 由 `ui/ui_window.py::load_engine` 按 `here` 传入的路径查找。

架构惯例：**界面与引擎分离**；引擎只做计算与 IO、可被命令行直接调用；界面只做调度与展示。
后台任务统一走 `Runner`（子线程投递队列 → 主线程 `after()` 轮询更新 UI）。

---

## 3. 运行环境与依赖

| 项 | 要求 |
|---|---|
| Python | **3.9+**（实测 3.13.12 可用） |
| openpyxl | **3.1.5**（必须，读写 xlsx 且要保留公式） |
| tkinter | 必须能**真正创建窗口**（不是只 import 成功） |

```bash
pip install openpyxl==3.1.5
```

⚠️ **Windows 上的 tkinter 陷阱**：官方 Python for Windows 自带 tkinter，一般没问题；
但如果用 `pythonw.exe` 启动会没有控制台（日志看不到），建议用 `python.exe`。
若报 `No module named tkinter`，说明装的是 embeddable 版本，换标准安装版。

---

## 4. 需要改的 3 处（详解）

### 4.1 启动器：`.command` → `.bat`

macOS 的 `启动.command` 做三件事：`cd` 到脚本目录 → 探测一个**能真正开窗**的解释器 → 启动 `main.py`（可带拖入的路径作 `argv[1]`）。

Windows 对应写法（`启动.bat`）：

```bat
@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 探测能真正开窗的解释器（tkinter + openpyxl 都要有）
set "PY="
for %%P in (python python3 py) do (
    if not defined PY (
        %%P -c "import tkinter,openpyxl;r=__import__('tkinter').Tk();r.withdraw();r.destroy()" >nul 2>&1 && set "PY=%%P"
    )
)
if not defined PY (
    echo 未找到可用的 Python（需 tkinter + openpyxl）
    pause & exit /b 1
)

"%PY%" main.py %1
pause
```

要点：
- **`%1` 是拖到 .bat 图标上的文件夹路径**，`main.py` 已经支持 `sys.argv[1]` 作为预置路径，直接传即可。
- `chcp 65001` 解决中文输出乱码；若仍有乱码，设 `set PYTHONIOENCODING=utf-8`。
- `cd /d "%~dp0"` 保证工作目录是脚本所在目录（引擎按 `os.path.dirname(__file__)` 定位，其实不强依赖，但保持稳妥）。
- Windows **没有可执行位**概念，不需要 `chmod 755`。

### 4.2 `ui/ui_window.py`：路径与"打开"

现有代码：

```python
def downloads_dir():                      # macOS: ~/Downloads 或 ~/下载
    home = os.path.expanduser('~')
    for c in (os.path.join(home,'Downloads'), os.path.join(home,'下载')):
        if os.path.isdir(c): return c
    return home

def open_path(p):                         # macOS: open
    if IS_MAC: subprocess.run(['open', p], check=False)
    ...
def reveal_path(p):                       # macOS: open -R
    if IS_MAC: subprocess.run(['open','-R', p], check=False)
```

Windows 改法：

```python
def downloads_dir():
    # Windows 优先用注册表里的真实「下载」目录（用户可能改过位置），拿不到再退回 ~/Downloads
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
              r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders')
        val, _ = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')
        winreg.CloseKey(key)
        val = os.path.expandvars(val)
        if os.path.isdir(val):
            return val
    except Exception:
        pass
    return os.path.join(os.path.expanduser('~'), 'Downloads')

def open_path(p):
    if IS_WIN:
        os.startfile(p)                   # 用默认程序打开
        return True
    ...

def reveal_path(p):
    if IS_WIN:
        subprocess.run(['explorer', '/select,', os.path.normpath(p)], check=False)
        return True
    ...
```

`IS_WIN = platform.system() == 'Windows'` 已经在文件里定义好了，直接复用。

### 4.3 `ui/ui_design.py`：字体

```python
class Fonts(object):
    def __init__(self):
        avail = set(tkfont.families())
        # macOS: PingFang SC / Menlo
        # Windows: Microsoft YaHei UI / Consolas
        fam  = 'Microsoft YaHei UI' if 'Microsoft YaHei UI' in avail else (
               'Microsoft YaHei' if 'Microsoft YaHei' in avail else 'TkDefaultFont')
        mono = 'Consolas' if 'Consolas' in avail else 'TkFixedFont'
```

建议写成"按平台选候选、再按可用字体回退"，这样一份代码两边都能跑。

---

## 5. 其他平台差异（不是必须改，但要知道）

| 事项 | 说明 |
|---|---|
| **Excel 占用文件** | Windows 上用 Excel 打开着的 xlsx 会被加锁，写入会抛 `PermissionError`（macOS 比较宽松）。**务必在文档/界面提示用户先关闭所有 Excel** |
| **临时锁文件** | macOS 会生成 `._xxx`（AppleDouble）和 `.~xxx`；Windows 主要生成 `~$xxx`。引擎统一排除 `('~$', '.~', '._')`，**这份排除逻辑保留即可，Windows 上同样安全** |
| **编码** | 所有 `open(..., encoding='utf-8')` 都已是显式指定，Windows 下不会因默认 GBK 出错。写 md/xlsx 处同样已指定 |
| **中文文件名/路径** | Python 3 + Windows 一般无问题；保持用 `os.path.join`，不要手工拼分隔符 |
| **权限位** | Windows 无 `chmod`，脚本不需要 755 |
| **拖拽启动** | macOS 拖到 `.command`；Windows 拖到 `.bat`，路径走 `%1` |

---

## 6. 为什么业务代码不用改

- 全部是标准库 + openpyxl，无 shell 调用、无平台 API（除上面 3 处）。
- 路径全部用 `os.path.join` / `os.path.expanduser`。
- 文件读写都用 `utf-8` 显式声明。
- `load_engine()` 按 `os.path.dirname(__file__)` 定位同目录引擎，换系统照样能找到。

⚠️ 唯一要注意：如果你**新增**了任何写文件的代码，请照抄现有写法显式带 `encoding='utf-8'`。

---

## 7. 业务规则全集（移植后请照此校验）

### ① 两表核对（`verify.py`）
- 输入：年级文件夹下**每班一个子文件夹**，内含 `综合素质测评详情表.xlsx` + `综合素质测评表.xlsx`
- 逐人比对打分项与总分，问题分三级：**必须处理 / 需确认 / 提示**
- 上限/扣分规则写在 `verify.py` 的 `CAP` 与 `DEDUCT_ITEMS`
- **只读源表，绝不写回**

### ② 挂科/附加分筛查（`挂科引擎.py`）
- **必修课** = 该课程列有成绩的人数 ≥ 班级实际人数 **90%**；零星列（重修/补修/智慧树公选）不算
- **挂科** = 必修列中 `<60` / `不及格` / `缺考·缓考·取消资格` 等未知文字
- 成绩标记：`60$`=补考、`60*`=重修，**剥离标记后按数值判，60 分算通过**
- **附加分** = 必修列**数字**成绩 `== 100`（五档制的「优」不算；只列资格，不算具体分值）
- 标黄：挂科学生**整行**（A 到最后一列，含汇总排名区）填标准黄 `FFFFFF00`
- **「人」与「门次」是两回事**：一人可挂多门，汇报时必须分开统计
- 跨年级跟班生（混入其他年级学号）**照常参与判定，不要过滤**
- 必修判定分母用**实际数据行数**，不用表头写的"班级人数"
- 复跑安全：产出写在目标目录内时，必须排除产出本身与锁文件，否则会被自己扫进来（自噬）

### ③ 回填综测总表（`回填引擎.py`）
- 综测总表结构：**第 1 行大标题、第 2 行表头、数据自第 3 行起**
  `A=姓名 B=学号 C=平均学分绩点 D=绩点折算 E(5)=附加分 F=学业成绩(=D+E) … J=综合测评成绩 K(11)=是否挂科 L=排名 M=签字`
- **是否挂科** ← 「挂科汇总」学号集合：命中填 `是`，否则 `否`
- **附加分** ← 「附加分汇总」的 `100分门数`：1 门=1，几门加几分，**没有必须填 0**（E 列参与 `F=D+E` 公式，不能留空）
- **学号归一化**：统一 12 位文本；比对前 `strip()` 并去掉 Excel 文本前缀撇号（`'202420010301` 这种）
- **身份双重确认**：学号 + 班级
  - 学号结构 = 年级4位 + 专业码4位 + 班号2位 + 序号2位
  - 班级名 = 专业名 + 年级后2位 + 专业码前2位 + 班号（例：`202520010501` → `计算机252005`）
  - 专业码映射：`2001`计算机 / `2004`软件工程 / `2005`物联网 / `2006`智能科学 / `2007`大数据
  - 专业码不在映射内（跟班生 `1608`/`1609`/`1802`）→ 跳过校验，不误报
- **年级对齐**（重要）：源表常含多个年级。回填某年级目录时，先用该目录学生学号的**年级众数**锁定年级，只取源表同年级名单。否则别的年级学生会被全部误报成"名单有·总表无"
- **跨年级跟班生**：学号年级 ≠ 名单班级年级 → 属正常，**跳过班级比对**，不算错
- **总表缺人**（名单有、表里没这行）→ 降级为警告，不算填充失败
- 写入前**一律先备份**到 `<目录>/_备份_修改前_时间戳/`；一律按名单覆盖

---

## 8. 引擎接口契约（便于重写或校验）

### `回填引擎.py`（③，CLI 可直接跑）
```bash
python 回填引擎.py <总表目录> <挂科与附加分表.xlsx> [--dry-run] [--no-backup]
```
主要函数：
```python
load_source(path)                 -> (fail_map, add_map, info)
    # fail_map: {学号: {'g':年级, 'k':班级}}
    # add_map : {学号: (门数, 年级, 班级)}      门数已按人聚合，不要再累加行数
collect_targets(root)             -> [(标签, [xlsx...]), ...]   # 自适应单/多年级
find_header_row(ws)               -> int|None    # 前 6 行内找含 姓名/学号/是否挂科 的行
locate_cols(ws, hdr)              -> {'name','sid','add','fail'}
scan_rows(path)                   -> [(sheet, 行号, 姓名, 学号, 附加分, 是否挂科)]
prepare(root, src_path)           -> (src, targets, plans)      # plans 已按年级过滤好
phase_dryrun(plans, src)          -> {'per_file','issues','total'}
phase_apply(plans, src, backup)   -> (核验是否通过, 备份目录列表)
phase_verify(plans, src)          -> bool
write_report(out_dir, plans, src, res, backups, ok) -> 报告路径
```

### `挂科引擎.py`（②）
```python
collect_xlsx(root)                -> [文件路径]
phase_dryrun(files)               -> [每班结果 dict]
phase_apply(results)              -> 写标黄
phase_verify(results)             -> bool
build_workbook(results, out, meta, include_course_audit=True) -> stat
OUT_NAME = '挂科与附加分表.xlsx'
```
产出 7 个 Sheet：`挂科汇总` / `挂科明细` / `附加分汇总` / `附加分明细` / `班级统计` / `课程认定明细` / `规则说明`
（其中「规则说明」页写明了下游字段映射，是②③之间的契约文档）

### `verify.py`（①）
界面调 `verify.verify_class` / `merge_common_issues` / `build_report`（**不要复制规则，直接调用**）。

---

## 9. 验收基线（移植后必须能对上）

### ② 全量基线（54 个班）

| 口径 | 数值 |
|---|---:|
| 挂科**人数** | 550 |
| 挂科**门次** | 1042 |
| 附加分人数 | 29 |
| 100 分门次 | 30 |

按年级：2023 级 64人/87门次 · 2024 级 181人/306门次 · 2025 级 305人/649门次
交叉核验必须 **54/54 全部一致**。

### ③ 各年级总表基线

| 年级 | 学生行 | 是 | 否 | 附加分 | 已知异常 |
|---|---:|---:|---:|---|---|
| 23 级 | 592 | 62 | 530 | 9 人/9 门次 | 4 个学号位数异常 + 2 人总表缺 |
| 24 级 | 667 | 181 | 486 | 8 人/8 门次 | 16 条跨年级跟班生（正常） |
| 25 级 | 751 | 302 | 449 | 12 人/13 门次 | 3 人总表缺 + 2 学号异常 + 1 身份存疑 |

（25 级附加分 13 门次是因为张静茹 `202520010727` 一人 2 门。）

---

## 10. 自测方法（推荐照做）

用**无头方式驱动 Tk**：加载模块 → `tk.Tk()` → 建 `Toplevel` → `build(...)` → `invoke()` 按钮 → 循环 `update()` 等日志出现标记。

```python
import tkinter as tk, tkinter.messagebox as mb
mb.askyesno = lambda *a, **k: True          # 必须 stub，否则弹窗阻塞
mb.showinfo = mb.showwarning = mb.showerror = lambda *a, **k: None
root = tk.Tk(); root.withdraw()
win = tk.Toplevel(root); win.withdraw()
mod.open_path = mod.reveal_path = lambda p: None   # 别真的打开资源管理器
mod.build(win, Fonts(), out_root, preset)
btn.invoke()                                 # 用控件的 text 前缀找按钮
while 未超时: root.update(); win.update(); 判断日志文本含标记
```

⚠️ 三个必须记住的测试坑：
1. **每个页面必须独立进程跑**。同一进程里连续创建多个 `tk.Tk()` 会崩（macOS 实测），Windows 同样建议分开跑。
2. **等待标记必须选日志里真实出现的字符串**。例如①的"总人数"是 summary、不进日志，要用"报告："；③要用"✔ 报告："。（本次在 macOS 上就因为这个白等了 300 秒）
3. **不要用 `| tail` 看输出**（管道缓冲，看不到进度，会误判卡死）；重定向到文件，或用 `python -u`。

测试数据建议：把少量班级复制到临时目录再跑，**不要拿真实数据做写入类测试**；② 的标黄和③的回填都会改文件。

---

## 11. 已知坑清单（macOS 上踩过的，Windows 也值得留神）

1. **学号/姓名列方向**：班级成绩表是 `A=学号、B=姓名`；**综测总表是 `A=姓名、B=学号`，正好相反**。
2. **学号带 Excel 文本前缀撇号**：如 `'202420010301`，必须 `lstrip("'")`，否则整段失配（24 级计算机表 242003 班整段如此，其中还包含 2 名附加分学生）。
3. **复跑自噬**：产出写在目标目录内，第二次跑会被自己扫进来；macOS 还会生成 `.~xxx.xlsx`（165 字节，不是 zip）导致 `BadZipFile` 整批崩。**必须排除 `~$` / `.~` / `._` / 产出本身**，且单文件 load 要 try/except 跳过。
4. **附加分名单有缺班级列的重复行**：解析要取「门数列最大值」（`max`），**不要按行 +1 累加**，否则多门学生会被少算。
5. **核验脚本要把"名单有但表中无""学号位数异常但姓名不在名单"降级为警告**，否则会把总表缺人/源数据错误误判成填充失败。
6. **Tk 的 `padx/pady` 只接受 1 或 2 个值**，写三元组 `pady=(6,0,0)` 直接 TclError。
7. **子线程不能碰 Tk 变量**：所有界面取值要在主线程快照成 dict 再传给 worker，否则 `main thread is not in main loop`。
8. **`sync_opt` 类闭包**：`command=sync_opt` 里引用的函数必须先定义，否则 `NameError: cannot access free variable`。
9. **`py_compile` 会写 `__pycache__`**，做纯语法检查请用 `compile(open(f).read(), f, 'exec')`。
10. **macOS 不能用系统自带 python3**：Tcl/Tk 8.5 能 import 但创建窗口必崩，必须做"真实开窗探测"。Windows 一般自带 tkinter 没问题，但仍建议探测。

---

## 12. 交接检查清单（Windows 适配做完后逐条打勾）

- [ ] `启动.bat` 能双击打开首页，且拖文件夹到图标上能带进路径
- [ ] 三个任务页都能打开，中文不乱码
- [ ] ① 用 4 个班的样本跑通，产出核验报告 xlsx，源表零改动
- [ ] ② 用 3 个班样本跑通：标黄行数正确、交叉核验 N/N 一致、产出 7 Sheet
- [ ] ③ 用总表副本跑通：备份生成、写入行数正确、**核验通过**、报告生成
- [ ] ② → ③ 联动：③ 能自动找到②产出的《挂科与附加分表.xlsx》
- [ ] 对照第 9 节基线数据，结果一致
- [ ] 提示用户：运行前关闭所有打开的 Excel

---

## 13. 已封装为 Windows 单文件 exe（PyInstaller）

> 本程序已在本机 Windows 上用 PyInstaller 打成**单个 `综测核算程序.exe`**（约 20 MB，双击即用，无需安装 Python/openpyxl）。

### 环境坑（重要）
- 本机**托管 Python 3.13.12 不含 tkinter**（精简版，无 tcl/tk），不能用来打 GUI 包。
- 改用**商店版 Python 3.11.9**（自带 tkinter / Tcl-Tk 8.6.12 + openpyxl 3.1.5）建 venv 后再打包。
- 若你回自己电脑打包：必须确保打包用的 Python **真正带 tkinter**（能 `Tk().withdraw()` 成功），否则打出的 exe 启动即崩。

### 构建命令
```bash
# 1) 用带 tkinter 的 Python 建 venv 并装依赖
py -m venv buildenv
buildenv\Scripts\pip install pyinstaller openpyxl==3.1.5

# 2) 生成单文件、无控制台窗口的 spec（已就绪：综测核算程序.spec）
#    —— 三个引擎经 load_engine 按文件路径加载，已作为 datas 内嵌；
#       openpyxl 已设为 hiddenimport。
buildenv\Scripts\pyi-makespec --onefile --windowed --name 综测核算程序 main.py
#    （然后手动在 spec 的 Analysis 里加 datas=[('verify.py','.'),('挂科引擎.py','.'),('回填引擎.py','.')] 与 hiddenimports=['openpyxl']）

# 3) 构建
buildenv\Scripts\pyinstaller 综测核算程序.spec
# 产出：dist\综测核算程序.exe
```

### 验证结论（本机）
- ✅ 三个引擎（含中文名 `挂科引擎.py`/`回填引擎.py`）已嵌入 exe，且 `load_engine` 可取。
- ✅ 窗口版 exe 实测可启动、首页 GUI 正常创建。
- ⚠️ 单文件 exe 每次运行会解包到 `%TEMP%\_MEIxxxx` 并多起一个进程，属正常现象。
- ⚠️ 目标机器无需装 Python；但若缺 VC++ 运行库，PyInstaller 一般已把 `vcruntime140.dll` 打进 exe，通常可直接跑。

### 13.x 数字签名 / 版本 / 安装包（新增）

本程序现内置三件套，封装时即可产出带版本、带签名的 exe，并配套安装包：

- **版本资源** `version_info.txt`（作者=程震宇、v1.0.0.0），由 spec 的 `version=` 嵌入 exe。
- **签名脚本** `sign_exe.py`：默认用 PowerShell 生成 CN=程震宇 的自签名证书签名；
  设 `SIGN_PFX` / `SIGN_PWD`（及可选 `SIGN_TIMESTAMP`）即用正规 CA 证书签名。
- **一键构建** `build.py`：`PyInstaller 打包 → sign_exe 签名`，等价于
  `pyinstaller 综测核算程序.spec` 之后再签名。
- **安装包** `installer.iss`：Inno Setup 脚本，装 Inno Setup 后 `ISCC installer.iss`
  生成 `综测核算程序_Setup.exe`（开始菜单 + 桌面快捷方式 + 卸载）。

> 自签名证书签名**真实有效**，但 Windows 会提示"发布者未验证"（未经 CA 信任）。
> 若要"可信发布者=程震宇"且不触发 SmartScreen，必须用正规代码签名证书（.pfx）。

### 交付后仍需你这边做的
- 用真实样例班级对照第 9 节基线（挂科 550 人 / 1042 门次、附加分 29 人 / 30 门次）跑一遍三任务，确认结果一致。
- 提醒使用者在运行前关闭所有打开的 Excel（Windows 会锁文件）。
