#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 班级成绩处理 · 生成《挂科与附加分表.xlsx》   v3.0
 太原科技大学 计算机科学与技术学院
=============================================================================

用法：
    python3 生成挂科附加分表.py [年级文件夹 or 学期根目录] [选项]

选项：
    --no-highlight   不写标黄，只生成表格（源文件只读不修改）
    -y / --yes       跳过确认，直接执行

不传路径则交互式询问（支持把文件夹拖进来后回车）。
目标可以是一个年级目录，也可以是包含多个年级子目录的学期根目录。

-----------------------------------------------------------------------------
产出：目标目录下《挂科与附加分表.xlsx》，含 7 个 Sheet
-----------------------------------------------------------------------------
  挂科汇总        一人一行（只有挂科学生）   ← 下游回填「是否挂科」用这张
  挂科明细        一门一行（含课程/列标/原始成绩）
  附加分汇总      一人一行（只有拿满100的人）← 下游回填「附加分」用这张
  附加分明细      一门一行
  班级统计        每班一览（人数/必修数/挂科数/附加分数）
  课程认定明细    审计用：每一列为什么算必修、为什么算选修
  规则说明        判定规则 + 运行元信息 + 下游字段映射说明

-----------------------------------------------------------------------------
判定规则（与历史基数完全一致的口径，勿擅自改动）
-----------------------------------------------------------------------------
  必修课  = 该课程列「有成绩人数 ÷ 实际数据行数 ≥ 90%」
            其余零星列（重修/补修/智慧树公选）不参与挂科与附加分判定
  挂科    必修列中：数字 <60 / 「不及格」/ 缺考·缓考·0分·取消资格等未知文字
            「60$」补考、「60*」重修 剥离标记按数值判，60 分算通过
            空白不参与判定
  汇总区  从「算术平均分」列起到最后一列，不参与判定
  附加分  必修列的**数字**成绩 == 100
            五档制的「优」不算；选修/零星列不算
            只列资格，不计算具体加分值（分值由学院核定）

  填色    挂科学生整行标准黄 FFFFFF00（A 列 ~ 最后一列，含汇总排名区）

-----------------------------------------------------------------------------
已知坑位（都已在本程序中处理，改代码前务必先读）
-----------------------------------------------------------------------------
  1. 班级成绩表 A 列=学号、B 列=姓名（第6行有表头）——与综测总表 A=姓名/B=学号
     正好相反，切勿混用。
  2. 学号可能存成 int / float / str，甚至带 Excel 文本前缀撇号。输出到 xlsx 时
     一律写成 12 位文本，并在规则页提示下游比对前做归一化。
  3. 同一个学生在同一学期可能有多门 100 分（如张静茹 2 门）。汇总 Sheet 已按人
     聚合好门数，下游直接取值，不要自己再累加行数。
  4. 必修判定的分母用「实际数据行数」而非表头的「班级人数」——历史基数就是这么
     算出来的，改分母会造成必修列认定漂移。
  5. 班里常混入其他年级学号（降级重修跟班生），属正常，照常参与判定。
=============================================================================
"""

import os
import re
import sys
import time

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write(
        "\n[错误] 缺少依赖 openpyxl。\n"
        "请先执行：  pip install openpyxl\n\n"
    )
    sys.exit(1)


# ------------------------------------------------------------------ 常量
YELLOW = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
REQUIRED_RATIO = 0.9        # 必修课覆盖率阈值 —— 历史口径，勿改
PERFECT_SCORE = 100         # 附加分认定分数线（只认数字 100）
PASS_TEXT = {'优', '良', '中', '及格', '合格', '通过', '优秀', '良好', '中等'}
FAIL_TEXT = {'不及格', '不合格', '未通过', '不通过'}
ID_LEN = 12                 # 学号位数

OUT_NAME = '挂科与附加分表.xlsx'

HFILL = PatternFill(start_color='FFD9E1F2', end_color='FFD9E1F2', fill_type='solid')
HFONT = Font(bold=True)

if sys.platform == 'win32':
    os.system('color')
    C = {'R': '', 'G': '', 'Y': '', 'B': '', 'D': '', 'E': ''}
else:
    C = {'R': '\033[31m', 'G': '\033[32m', 'Y': '\033[33m',
         'B': '\033[1m', 'D': '\033[2m', 'E': '\033[0m'}


def hr(ch='─'):
    try:
        w = min(os.get_terminal_size().columns, 96)
    except Exception:
        w = 78
    print(ch * w)


def title(t):
    print()
    hr('═')
    print(' ' + t)
    hr('═')


# ------------------------------------------------------------------ 路径
def clean_path(raw):
    if raw is None:
        return ''
    s = raw.strip().strip('\r\n')
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'", '`'):
        s = s[1:-1]
    s = s.replace('\\ ', ' ').replace('\\(', '(').replace('\\)', ')')
    s = s.replace("\\'", "'").replace('\\"', '"')
    s = os.path.expanduser(s)
    try:
        s = os.path.realpath(s)
    except Exception:
        pass
    return s


def ask_path():
    print()
    print('  提示：可直接把文件夹拖进来再按回车，也可粘贴完整路径。')
    print('  ' + '─' * 60)
    while True:
        try:
            raw = input('  ' + C['B'] + '请拖入 / 输入 成绩文件夹路径' + C['E'] + '  › ').strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None
        if raw.lower() in ('q', 'quit', 'exit'):
            return None
        if not raw:
            continue
        p = clean_path(raw)
        if os.path.isdir(p):
            return p
        print('  ' + C['R'] + '✗ 路径不存在或不是文件夹：' + C['E'] + p)


def collect_xlsx(target):
    """扫描目标下所有班级成绩表。

    必须排除的几类垃圾（否则复跑会崩 / 会把自己吃掉）：
      ~$xxx   Excel 打开时的临时锁文件
      .~xxx   macOS 打开文件时生成的隐藏锁文件（不是 zip，load 会抛 BadZipFile）
      ._xxx   macOS 在非 HFS 卷生成的 AppleDouble 资源叉
      产出本身  上一轮生成的《挂科与附加分表.xlsx》会被自己扫进来
    """
    junk_prefix = ('~$', '.~', '._')
    out = []
    for root, dirs, files in os.walk(target):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for f in files:
            if not f.lower().endswith('.xlsx'):
                continue
            if f.startswith(junk_prefix):
                continue
            if f == OUT_NAME:          # 自己的产出不能当输入，否则复跑污染
                continue
            out.append(os.path.join(root, f))
    return sorted(out)


# ------------------------------------------------------------------ 判定核心
def norm_id(v):
    """学号归一化：int/float/str 统一；剥离 Excel 文本前缀撇号。"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().lstrip("'").strip()


def is_student_row(v):
    s = norm_id(v)
    return bool(s) and s.isdigit() and len(s) >= 8


def parse_score(v):
    """返回 ('empty' | 'pass' | 'fail', 数值或原字符串)"""
    if v is None:
        return 'empty', None
    s = str(v).strip()
    if s == '':
        return 'empty', None
    m = re.match(r'^(\d+(?:\.\d+)?)\s*[$*]?$', s)
    if m:
        num = float(m.group(1))
        return ('pass' if num >= 60 else 'fail'), num
    if s in PASS_TEXT:
        return 'pass', s
    if s in FAIL_TEXT:
        return 'fail', s
    return 'fail', s          # 缺考 / 缓考 / 取消资格 等未知文字 → 挂科


def find_summary_col(ws):
    for c in range(1, ws.max_column + 1):
        v = ws.cell(5, c).value
        if v and '算术平均分' in str(v):
            return c
    return None


def parse_meta(ws):
    """从表头第 2 行解析 年级 / 班级 / 班级人数"""
    txt = str(ws.cell(2, 1).value or '')
    g = re.search(r'年级[：:]\s*(\d{4})', txt)
    k = re.search(r'班级[：:]\s*([^\s(（]+)', txt)
    n = re.search(r'班级人数[：:]\s*(\d+)', txt)
    return (g.group(1) if g else '',
            k.group(1) if k else '',
            int(n.group(1)) if n else None)


def scan_sheet(ws):
    """扫描单个班级表，返回识别结果字典"""
    grade, klass, roster = parse_meta(ws)

    sum_col = find_summary_col(ws)
    if sum_col is None:
        sum_col = ws.max_column + 1        # 兜底：没找到就看完整表宽
    last_course = sum_col - 1

    rows = [r for r in range(8, ws.max_row + 1)
            if is_student_row(ws.cell(r, 1).value)]
    n = len(rows)

    # ---- 逐列判定必修 / 选修（含审计信息）
    required, colinfo = set(), []
    for c in range(4, last_course + 1):
        cnt = sum(1 for r in rows if parse_score(ws.cell(r, c).value)[0] != 'empty')
        name = str(ws.cell(5, c).value or '').strip()
        ratio = round(cnt / n * 100, 1) if n else 0.0
        is_req = cnt >= n * REQUIRED_RATIO
        if is_req:
            required.add(c)
        colinfo.append((get_column_letter(c), name, cnt, ratio,
                        '必修' if is_req else '选修/零星'))

    # ---- 挂科 / 附加分
    fails, perfect, fail_rows = [], [], set()
    for r in rows:
        sid = norm_id(ws.cell(r, 1).value)
        name = str(ws.cell(r, 2).value or '').strip()
        bad, per = [], []
        for c in sorted(required):
            kind, val = parse_score(ws.cell(r, c).value)
            course = str(ws.cell(5, c).value or '').strip()
            raw = ws.cell(r, c).value
            if kind == 'fail':
                bad.append((get_column_letter(c), course, raw))
            elif (kind == 'pass' and isinstance(val, (int, float))
                  and val == PERFECT_SCORE):
                per.append((get_column_letter(c), course, raw))
        if bad:
            fails.append((sid, name, bad))
            fail_rows.add(r)
        if per:
            perfect.append((sid, name, per))

    return {
        'grade': grade, 'klass': klass, 'roster': roster,
        'rows': rows, 'n': n,
        'required': required, 'colinfo': colinfo,
        'sum_col': sum_col,
        'fails': fails, 'perfect': perfect, 'fail_rows': fail_rows,
    }


# ------------------------------------------------------------------ 阶段
def phase_scan(target, files):
    title('[1/6] 扫描文件夹')
    print('  目标：' + C['B'] + target + C['E'])
    print('  找到 ' + C['B'] + '%d' % len(files) + C['E'] + ' 个班级成绩表')
    return files


def phase_dryrun(files):
    title('[2/6] 空跑判定（不修改任何文件）')
    results, skipped = [], []
    for p in files:
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
        except Exception as e:
            # 打不开的文件（临时锁、损坏文件）跳过并告警，不能让整个批次崩掉
            skipped.append((os.path.basename(p), '无法打开：%s' % e))
            continue
        ws = wb[wb.sheetnames[0]]
        d = scan_sheet(ws)
        wb.close()
        if d['n'] == 0:
            # 没有学生行的表不参与统计（也是防止非成绩表被误纳入的兜底）
            skipped.append((os.path.basename(p), '未识别到学生行，已跳过'))
            continue
        d['path'] = p
        d['name'] = os.path.basename(p).replace('班级成绩(', '').replace(').xlsx', '')
        results.append(d)
    results.sort(key=lambda x: (x['grade'], x['name']))

    if skipped:
        print('  ' + C['Y'] + '! 已跳过 %d 个非成绩文件：' % len(skipped) + C['E'])
        for nm, why in skipped:
            print('      · %s（%s）' % (nm, why))
        print()

    cur = None
    for d in results:
        if d['grade'] != cur:
            cur = d['grade']
            print('\n  ── %s级 ──' % cur)
        f = (C['Y'] + '挂科 %2d' % len(d['fails']) + C['E']) if d['fails'] \
            else (C['G'] + '无挂科' + C['E'])
        b = (C['G'] + '%d人100分' % len(d['perfect']) + C['E']) if d['perfect'] \
            else (C['D'] + '无100分' + C['E'])
        print('    %-18s %3d人 必修%2d门  %s   %s'
              % (d['name'], d['n'], len(d['required']), f, b))

    ts = sum(d['n'] for d in results)
    tf = sum(len(d['fails']) for d in results)
    tfc = sum(len(x[2]) for d in results for x in d['fails'])
    tp = sum(len(d['perfect']) for d in results)
    tpc = sum(len(x[2]) for d in results for x in d['perfect'])
    hr()
    print('  ' + C['B'] + '合计：%d 个班 · %d 名学生 · 挂科 %d 人/%d 门次 · '
          '附加分 %d 人/%d 门次' % (len(results), ts, tf, tfc, tp, tpc) + C['E'])
    return results


def phase_apply(results):
    title('[3/6] 写入标黄')
    for i, d in enumerate(results, 1):
        wb = openpyxl.load_workbook(d['path'])
        ws = wb[wb.sheetnames[0]]
        last_col = ws.max_column
        for r in d['rows']:
            for c in range(1, last_col + 1):
                cell = ws.cell(r, c)
                if r in d['fail_rows']:
                    cell.fill = YELLOW
                else:
                    fl = cell.fill
                    if fl is not None and fl.patternType == 'solid' \
                            and fl.fgColor.rgb == 'FFFFFF00':
                        cell.fill = PatternFill(fill_type=None)
        wb.save(d['path'])
        wb.close()
        print('  ✓ (%d/%d) %-18s 标黄 %2d 行'
              % (i, len(results), d['name'], len(d['fail_rows'])))


def phase_verify(results):
    title('[4/6] 交叉核验')
    print('  独立重算期望结果，与文件内实际标黄逐行比对…\n')
    ok, bad = 0, []
    for d in results:
        wb = openpyxl.load_workbook(d['path'])
        ws = wb[wb.sheetnames[0]]
        last_col = ws.max_column
        expect = set()
        for r in d['rows']:
            for c in sorted(d['required']):
                if parse_score(ws.cell(r, c).value)[0] == 'fail':
                    expect.add(r)
                    break
        actual, partial = set(), []
        for r in d['rows']:
            cols = {c for c in range(1, last_col + 1)
                    if ws.cell(r, c).fill and ws.cell(r, c).fill.patternType == 'solid'
                    and ws.cell(r, c).fill.fgColor.rgb == 'FFFFFF00'}
            if not cols:
                continue
            if cols == set(range(1, last_col + 1)):
                actual.add(r)
            else:
                partial.append(r)
        wb.close()
        if expect == actual and not partial:
            ok += 1
        else:
            bad.append('%s：期望%d 实际%d 多标%s 漏标%s 残缺%s'
                       % (d['name'], len(expect), len(actual),
                          sorted(actual - expect), sorted(expect - actual), partial))
    hr()
    if bad:
        for b in bad:
            print('  ' + C['R'] + '!! ' + b + C['E'])
        print('\n  ' + C['R'] + '核验未通过：%d / %d 一致' % (ok, len(results)) + C['E'])
        return False
    print('  ' + C['G'] + '✓ 核验通过：%d / %d 个文件全部一致（无多标/漏标/残缺）'
          % (ok, len(results)) + C['E'])
    return True


# ------------------------------------------------------------------ 输出 xlsx
def write_sheet(wb, name, headers, rows, widths=None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font = HFONT
        cell.fill = HFILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.freeze_panes = 'A2'
    if rows:
        ws.auto_filter.ref = 'A1:%s%d' % (get_column_letter(len(headers)), len(rows) + 1)
    # 学号列强制文本，防止 Excel 科学计数法 / 精度丢失
    if '学号' in headers:
        letter = get_column_letter(headers.index('学号') + 1)
        for r in range(2, len(rows) + 2):
            ws.cell(r, headers.index('学号') + 1).number_format = '@'
    for i, w in enumerate(widths or [14] * len(headers), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def build_workbook(results, out_path, meta, include_course_audit=True):
    """生成《挂科与附加分表.xlsx》。

    include_course_audit=False 时不输出「课程认定明细」Sheet（该表行数最多，
    纯业务用途时可以关掉以缩小文件体积）。GUI 里对应一个开关。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------- 挂科汇总（一人一行）
    rows = []
    for d in results:
        for sid, name, bad in d['fails']:
            rows.append([d['grade'], d['klass'], sid, name, len(bad),
                         '、'.join(x[1] or ('列' + x[0]) for x in bad)])
    write_sheet(wb, '挂科汇总',
                ['年级', '班级', '学号', '姓名', '挂科门数', '挂科课程'],
                rows, [8, 16, 16, 10, 10, 52])

    # ---------- 挂科明细（一门一行）
    rows = []
    for d in results:
        for sid, name, bad in d['fails']:
            for col, course, raw in bad:
                rows.append([d['grade'], d['klass'], sid, name,
                             course or ('列' + col), col, str(raw)])
    write_sheet(wb, '挂科明细',
                ['年级', '班级', '学号', '姓名', '挂科课程', '课程列', '原始成绩'],
                rows, [8, 16, 16, 10, 34, 8, 12])

    # ---------- 附加分汇总（一人一行，含聚合好的门数）
    rows = []
    for d in results:
        for sid, name, per in d['perfect']:
            rows.append([d['grade'], d['klass'], sid, name, len(per),
                         '、'.join(x[1] or ('列' + x[0]) for x in per)])
    write_sheet(wb, '附加分汇总',
                ['年级', '班级', '学号', '姓名', '100分门数', '100分课程'],
                rows, [8, 16, 16, 10, 12, 52])

    # ---------- 附加分明细
    rows = []
    for d in results:
        for sid, name, per in d['perfect']:
            for col, course, raw in per:
                rows.append([d['grade'], d['klass'], sid, name,
                             course or ('列' + col), col, str(raw)])
    write_sheet(wb, '附加分明细',
                ['年级', '班级', '学号', '姓名', '100分课程', '课程列', '成绩'],
                rows, [8, 16, 16, 10, 34, 8, 10])

    # ---------- 班级统计
    rows = []
    for d in results:
        rows.append([d['grade'], d['klass'], d['n'], len(d['required']),
                     len(d['colinfo']) - len(d['required']),
                     len(d['fails']),
                     sum(len(x[2]) for x in d['fails']),
                     len(d['perfect']),
                     sum(len(x[2]) for x in d['perfect'])])
    write_sheet(wb, '班级统计',
                ['年级', '班级', '在册人数', '必修门数', '选修列数',
                 '挂科人数', '挂科门次', '附加分人数', '100分门次'],
                rows, [8, 16, 10, 10, 10, 10, 10, 11, 11])

    # ---------- 课程认定明细（审计：每列为什么算必修/选修）
    # 开关关闭时整张 Sheet 都不创建（不是留一张空表）
    if include_course_audit:
        rows = []
        for d in results:
            for col, cname, cnt, ratio, verdict in d['colinfo']:
                rows.append([d['grade'], d['klass'], col, cname, cnt, d['n'], ratio, verdict])
        write_sheet(wb, '课程认定明细',
                    ['年级', '班级', '课程列', '课程名', '有成绩人数', '在册人数',
                     '覆盖率%', '认定结果'],
                    rows, [8, 16, 8, 36, 12, 10, 10, 12])

    # ---------- 规则说明
    lines = [
        ['项目', '内容'],
        ['生成时间', meta['time']],
        ['目标目录', meta['target']],
        ['处理班级', '%d 个' % len(results)],
        ['在册学生', '%d 名' % sum(d['n'] for d in results)],
        ['挂科人数', '%d 人' % sum(len(d['fails']) for d in results)],
        ['挂科门次', '%d 门次' % sum(len(x[2]) for d in results for x in d['fails'])],
        ['附加分人数', '%d 人' % sum(len(d['perfect']) for d in results)],
        ['100分门次', '%d 门次' % sum(len(x[2]) for d in results for x in d['perfect'])],
        ['标黄写入', meta['highlight']],
        ['交叉核验', meta['verify']],
        ['', ''],
        ['【判定规则】', ''],
        ['必修课认定', '课程列「有成绩人数 ÷ 在册人数 ≥ %.0f%%」' % (REQUIRED_RATIO * 100)],
        ['挂科认定', '必修列中：数字成绩 <60 / 「不及格」/ 缺考·缓考·0分·取消资格等未知文字'],
        ['补考重修标记', '「60$」补考、「60*」重修 → 剥离标记后按数值判，60 分算通过'],
        ['不参与判定', '空白单元格；汇总排名区（从「算术平均分」列起到最后一列）'],
        ['附加分认定', '必修列中的数字成绩 == 100'],
        ['附加分不计', '五档制的「优」、选修/重修/公选等零星列'],
        ['附加分说明', '本表只列资格，不计算具体加分值，分值由学院核定'],
        ['', ''],
        ['【下游填表字段映射】', ''],
        ['是否挂科', '综测总表 K 列(第11列) ← 取「挂科汇总」Sheet 的学号集合：命中填「是」，否则填「否」'],
        ['附加分', '综测总表 E 列(第5列) ← 取「附加分汇总」Sheet 的「100分门数」：无则填 0'],
        ['附加分(重要)', 'E 列参与 F 列公式 =D+E，未获资格也必须填 0，不能留空'],
        ['学号格式', '本表学号一律为 %d 位文本。比对前建议先做归一化（strip 并去掉可能的 Excel 文本前缀撇号）' % ID_LEN],
        ['多门100分', '「附加分汇总」已按人聚合好门数，下游直接取值，不要再累加行数'],
        ['', ''],
        ['【其他注意事项】', ''],
        ['列名陷阱', '班级成绩表是 A列=学号、B列=姓名；综测总表是 A列=姓名、B列=学号，方向相反，勿混用'],
        ['跨年级跟班生', '班里混入的其他年级学号（降级重修）属正常，照常参与判定'],
        ['同名双列', '同一课程可能出现两列（一列全员必修、一列几人重修），靠覆盖率自动分流，勿按课名去重'],
    ]
    ws = wb.create_sheet('规则说明')
    for ln in lines:
        ws.append(ln)
    for c in (1, 2):
        ws.cell(1, c).font = HFONT
        ws.cell(1, c).fill = HFILL
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 88
    for r in range(2, len(lines) + 1):
        ws.cell(r, 1).alignment = Alignment(vertical='top')
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
    ws.freeze_panes = 'A2'

    wb.save(out_path)
    return {
        'classes': len(results),
        'students': sum(d['n'] for d in results),
        'fail_stu': sum(len(d['fails']) for d in results),
        'fail_cnt': sum(len(x[2]) for d in results for x in d['fails']),
        'bonus_stu': sum(len(d['perfect']) for d in results),
        'bonus_cnt': sum(len(x[2]) for d in results for x in d['perfect']),
    }


# ------------------------------------------------------------------ 主流程
def run(target, do_highlight, auto_yes):
    files = collect_xlsx(target)
    phase_scan(target, files)
    if not files:
        print('  ' + C['R'] + '✗ 该目录下没有找到 .xlsx 文件' + C['E'])
        return 1

    results = phase_dryrun(files)
    if not results:
        return 1

    print()
    if do_highlight:
        print('  ' + C['D'] + '下一步：挂科学生整行填标黄 + 生成《挂科与附加分表.xlsx》' + C['E'])
    else:
        print('  ' + C['D'] + '下一步：只生成《挂科与附加分表.xlsx》（源文件保持只读）' + C['E'])

    if auto_yes:
        ans = 'y'
    else:
        try:
            ans = input('\n  ' + C['B'] + '确认执行？[y/N] ' + C['E']).strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            ans = 'n'
    if ans not in ('y', 'yes'):
        print('\n  已取消，未做任何修改。')
        return 0

    verified = '未执行（--no-highlight 模式）'
    if do_highlight:
        phase_apply(results)
        ok = phase_verify(results)
        verified = ('通过：%d/%d 全部一致' % (len(results), len(results))
                    if ok else '未通过')

    title('[5/6] 生成表格')
    out_dir = target if os.path.isdir(target) else os.path.dirname(target)
    out_path = os.path.join(out_dir, OUT_NAME)
    stat = build_workbook(results, out_path, {
        'time': time.strftime('%Y-%m-%d %H:%M'),
        'target': target,
        'highlight': '是（标准黄 FFFFFF00 整行）' if do_highlight else '否（--no-highlight）',
        'verify': verified,
    })

    title('[6/6] 完成')
    if do_highlight and not verified.startswith('通过'):
        print('  ' + C['R'] + '✗ 核验未通过，请检查' + C['E'])
    else:
        print('  ' + C['G'] + '✓ 全部完成' + C['E'])
    print('  · 处理班级：%d 个 · 在册学生：%d 名' % (stat['classes'], stat['students']))
    print('  · 挂科：%d 人 / %d 门次' % (stat['fail_stu'], stat['fail_cnt']))
    print('  · 附加分：%d 人 / %d 门次' % (stat['bonus_stu'], stat['bonus_cnt']))
    print('  · 输出文件：')
    print('      ' + out_path)
    hr('═')
    print()
    try:
        if sys.platform == 'darwin':
            os.system('open "%s"' % out_path)
        elif sys.platform == 'win32':
            os.startfile(out_path)
    except Exception:
        pass
    return 0


def main():
    args = sys.argv[1:]
    do_highlight = '--no-highlight' not in args
    auto_yes = ('-y' in args) or ('--yes' in args)
    positional = [a for a in args if not a.startswith('-')]

    title('班级成绩处理  v3.0   生成《挂科与附加分表.xlsx》')
    print('  太原科技大学 · 计算机科学与技术学院')
    print('  流程：扫描 → 空跑判定 → 写入标黄 → 交叉核验 → 生成表格')
    if not do_highlight:
        print('  ' + C['Y'] + '模式：--no-highlight（只出表，不改源文件）' + C['E'])

    target = None
    if positional:
        cand = clean_path(' '.join(positional))
        if os.path.isdir(cand):
            target = cand
            print('\n  已指定目标：' + C['B'] + target + C['E'])
        else:
            print('\n  ' + C['Y'] + '! 命令行路径无效，转为手动输入' + C['E'])
    if target is None:
        target = ask_path()
    if not target:
        print('\n  已退出。')
        return 0

    try:
        return run(target, do_highlight, auto_yes)
    except KeyboardInterrupt:
        print('\n\n  已中断。')
        return 130
    except Exception as e:
        print('\n  ' + C['R'] + '✗ 运行出错：' + C['E'] + str(e))
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
