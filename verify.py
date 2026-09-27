# -*- coding: utf-8 -*-
"""
综测两表核验脚本（macOS 版 v0.6，规则与 Windows v0.5 完全一致）
=================================================================
比对每个班级「综合素质测评详情表.xlsx」与「综合素质测评表.xlsx」：
总分及思想品德 / 文体活动 / 科技创新 / 志愿服务各分项是否对应。
规则：详情表文体活动栏不含测评表「体能测试(10分)」，故 测评表总分 − 详情表合计 = 10。

【安全约定 · 重要】
  本脚本对源 xlsx **仅做只读访问**：以 openpyxl 打开后一次性把单元格值抽成纯数据网格
  （Grid），随即关闭 workbook。调用方拿到的是纯数据（只有 .value），
  结构上没有任何可以写回源表的入口，绝不会改动源文件。
  唯一会执行写操作的是 `build_report()`：把结果导出到一个**新建**的报告簿并保存。

【v0.6 相对 v0.5 的改动（仅性能与跨平台，核验规则一字未改）】
  1) 读取改为「一次性读入内存网格」而非 read_only 随机访问，速度提升约 20 倍以上；
  2) 去掉对 read_only 的依赖，mac/win 行为完全一致。

用法：
  python verify.py --data "<年级文件夹>" [--term 学期标识] [--out 自定义报告路径]
"""

import argparse
import os
import glob
import re
from datetime import datetime
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# 配置
# ============================================================
DEFAULT_DATA = ""  # 命令行模式需显式传 --data；GUI 模式始终传路径，不依赖此默认值

# 评分细则上限（《学院综合素质测评细则2026.06修订》+ 审核手册）
CAP = {
    '政治素质': 15, '道德品质': 10, '组织纪律': 10, '其他': 10, '思想品德': 45,
    '体能测试': 10, '讲座活动': 10, '文体活动合计': 20,
    '科技竞赛': 10, '科技创新合计': 25,
    '志愿服务': 5, '社会实践': 5, '志愿服务合计': 10,
    '总计': 100,
}
# 扣分制项：无扣分材料即应得满分；低于满分需确认是否有扣分依据
DEDUCT_ITEMS = {'政治素质': 15, '道德品质': 10, '组织纪律': 10, '体能测试': 10}


# ============================================================
# 读取与安全封装（只读，绝不写回源文件）
# ============================================================
class _Cell:
    """只暴露 .value 的最小单元格，杜绝任何写回源表的可能。"""
    __slots__ = ('value',)

    def __init__(self, value=None):
        self.value = value


class Grid:
    """只读的二维数据网格，接口与 openpyxl 工作表保持一致（max_row/max_column/cell）。

    之所以不直接用 worksheet：openpyxl 的 read_only 模式对 ws.cell(r, c) 随机访问会
    反复解析整表，17 个班要跑十几分钟。改为一次性 iter_rows 抽成纯数据后关闭 workbook，
    访问变成 O(1)，同时因为只保留数值，结构上不可能写回源文件。
    """
    def __init__(self, rows):
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(r) for r in rows), default=0)

    def cell(self, row, column):
        try:
            return _Cell(self._rows[row - 1][column - 1])
        except (IndexError, TypeError):
            return _Cell(None)


def read_sheet(path):
    """只读读取一个 xlsx 的首个工作表，返回 (None, Grid)。
    第一个返回值保留为 None：源文件 workbook 在此已关闭，调用方拿不到它，也就无法保存。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.active
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return None, Grid(rows)


def find_header_row(ws):
    for r in range(1, min(ws.max_row, 12) + 1):
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(r, c).value
            if v and ('姓名' in str(v) or '学号' in str(v)):
                return r
    return None


def norm_section(t):
    if t is None:
        return None
    if '思想品德' in t:
        return 'sixiang'
    if '文体' in t:
        return 'wenti'
    if '科技' in t:
        return 'keji'
    if '志愿' in t or '社会' in t:
        return 'zhiyuan'
    if '总计' in t:
        return 'total'
    return None


def section_of_column(ws, header_row, col):
    for cc in range(col, 0, -1):
        v = ws.cell(header_row, cc).value
        if v is not None and str(v).strip() and str(v).strip() not in ('合计', '总计'):
            return str(v).strip()
    return None


def detect(ws):
    """按表头文字定位各列。返回 (header_row, name_col, xuehao_col, roles)。
    roles: 角色名 -> 列号，如 sixiang_heji / wenti_desc / keji_jingsai ..."""
    hr = find_header_row(ws)
    name_col = xuehao_col = None
    roles = {}
    for c in range(1, ws.max_column + 1):
        hv = ws.cell(hr, c).value
        if hv and '姓名' in str(hv):
            name_col = c
        if hv and '学号' in str(hv):
            xuehao_col = c
        texts = set()
        for rr in (hr, hr + 1):
            v = ws.cell(rr, c).value
            if v is not None and str(v).strip():
                texts.add(str(v).strip())
        if not texts:
            continue
        sec = norm_section(section_of_column(ws, hr, c))
        joined = ''.join(texts)
        if '合计' in texts:
            roles[sec + '_heji'] = c
        elif '总计' in texts or sec == 'total':
            roles['total'] = c
        elif '政治素质' in joined:
            roles['sixiang_zhengzhi'] = c
        elif '道德' in joined:
            roles['sixiang_pinde'] = c
        elif '纪律' in joined:
            roles['sixiang_jilv'] = c
        elif '其他' in joined:
            roles['sixiang_qita'] = c
        elif '体能' in joined:
            roles['wenti_tineng'] = c
        elif '讲座' in joined:
            roles['wenti_jiangzuo'] = c
        elif '竞赛' in joined:
            roles['keji_jingsai'] = c
        elif '论文' in joined:
            roles['keji_lunwen'] = c
        elif '专利' in joined:
            roles['keji_zhuanli'] = c
        elif '文学' in joined:
            roles['keji_wenxue'] = c
        elif '新闻' in joined:
            roles['keji_xinwen'] = c
        elif '资格' in joined:
            roles['keji_zige'] = c
        elif '志愿' in joined:
            roles['zhiyuan_fuwu'] = c
        elif '社会' in joined:
            roles['zhiyuan_shijian'] = c
        # 详情表的"描述列"(只写板块名、无子项)放在最后匹配，避免抢占测评表的子项列
        elif '文体活动' in joined:
            roles['wenti_desc'] = c
        elif '科技创新' in joined:
            roles['keji_desc'] = c
    return hr, name_col, xuehao_col, roles


def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ('', '无', 'None'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def approx(a, b, eps=0.01):
    if a is None or b is None:
        return False
    return abs(a - b) < eps


def vol_empty(v):
    """志愿服务单元格是否为空/无数据（空、'无'、0 均视为无数据）。"""
    return v is None or str(v).strip() in ('', '无', 'None') or num(v) <= 0.001


def norm_xh(v):
    return ''.join(ch for ch in str(v) if ch.isdigit()) if v is not None else ''


# 从详情表活动明细文本中反算加分合计，例如 "xx活动+0.3\nyy活动+0.5" -> 0.8
def sum_desc(txt, maxv=10.0):
    """反算明细加分（鲁棒版）。支持真实写法：
       - 半角/全角 +（＋）
       - 乘法 +0.3*3 / ＋1.2×4
       - 中文逗号 0，3、双点 0..3、句号 0。3 等笔误
       - 裸数字行（活动名后直接 0.3，无 + 号）
       - 编号误连 +0.38. （0.3 后接活动编号 8.）→ 截断为 0.3
       maxv 用于过滤年份(>=2000)与异常大数。
    """
    if txt is None:
        return None
    s = str(txt).strip()
    if s in ('', '无', 'None'):
        return 0.0
    # 归一化全角 / 笔误
    s = s.replace('＋', '+').replace('，', ',').replace('．', '.').replace('、', ',')
    s = s.replace('..', '.').replace('。', '.')
    total = 0.0
    found = 0
    tok = re.compile(r'[+]?\s*(\d+(?:\.\d+)?)\s*分?\s*(?:[×*xX]\s*(\d+))?')
    for raw in s.split('\n'):
        line = raw.strip()
        if not line:
            continue
        line_no = re.sub(r'^\s*\d+\s*[.、．]\s*', '', line)   # 去行首序号 N./N、
        for m in tok.finditer(line_no):
            v = float(m.group(1))
            if v <= 0 or v >= 2000:        # 排除年份 / 异常大数
                continue
            j = m.end()
            after = line_no[j:] if j < len(line_no) else ''
            # 编号误连：分值后紧跟 "." + 非数字（如 +0.38.纸短 → 0.3）
            if after.startswith('.') and len(after) > 1 and not after[1].isdigit():
                if '.' in m.group(1):
                    v = float(f"{int(v * 10) / 10:.1f}")   # 截断到一位小数
                else:
                    continue                            # 整数编号，跳过
            if v > maxv:
                continue
            mult = int(m.group(2)) if m.group(2) else 1
            total += v * mult
            found += 1
    if found == 0:
        # 兜底：整段就是一个数字或数字后紧跟左括号，如 "10"、"6（计算机协会主席…）"
        first = s.split('\n')[0].strip()
        m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*(?:[（(].*)?$', first)
        if m and float(m.group(1)) <= maxv:
            return float(m.group(1))
        return None
    return round(total, 2)


def load_file(path):
    """读取一个班的表为 {姓名: {role: 值, 'row': 行号}}。只读，不修改源文件。"""
    _, ws = read_sheet(path)
    hr, name_col, xuehao_col, roles = detect(ws)
    students = {}
    order = []
    for r in range(hr + 2, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if name is None or str(name).strip() == '':
            continue
        name = str(name).strip()
        rec = {'row': r, 'xuehao': ws.cell(r, xuehao_col).value}
        for role, col in roles.items():
            rec[role] = ws.cell(r, col).value
        students[name] = rec
        order.append(name)
    return roles, students, order


def detail_last_col_is_total(path):
    """判断详情表末列是否被填成了「总分」而非「志愿服务合计」。
    判据：多数学生 志愿描述为空 且 末列 == 思想+文体+科技 三者之和。"""
    try:
        _, ws = read_sheet(path)
    except Exception:
        return False
    hr, nc, xc, roles = detect(ws)
    zc = roles.get('zhiyuan_heji')
    dc = roles.get('zhiyuan_fuwu')
    if not zc:
        return False
    tot = 0
    hit = 0
    for r in range(hr + 2, ws.max_row + 1):
        nm = ws.cell(r, nc).value
        if nm is None or str(nm).strip() == '':
            continue
        sx = num(ws.cell(r, roles['sixiang_heji']).value) if 'sixiang_heji' in roles else 0.0
        if approx(sx, 0.0):
            sx = sum(num(ws.cell(r, roles[k]).value)
                     for k in ('sixiang_zhengzhi', 'sixiang_pinde', 'sixiang_jilv', 'sixiang_qita')
                     if k in roles)
        wt = num(ws.cell(r, roles['wenti_heji']).value) if 'wenti_heji' in roles else 0.0
        kj = num(ws.cell(r, roles['keji_heji']).value) if 'keji_heji' in roles else 0.0
        N = num(ws.cell(r, zc).value)
        desc = ws.cell(r, dc).value if dc else None
        empty_desc = desc is None or str(desc).strip() in ('', '无', 'None')
        tot += 1
        if empty_desc and abs(N - (sx + wt + kj)) < 0.01:
            hit += 1
    return tot > 0 and hit / tot >= 0.5


# ============================================================
# 问题记录
# ============================================================
def add(issues, part, typ, dv, ev, desc, advice, level, group=None, tpl=None):
    """记录一条问题。dv=详情表值, ev=测评表值, level=必须处理/需确认/提示
       group/tpl：整班共性问题用，同班同 group 达 5 条以上会合并成一行。"""
    diff = ''
    if isinstance(dv, (int, float)) and isinstance(ev, (int, float)):
        d = round(float(dv) - float(ev), 2)
        diff = d if abs(d) > 0.001 else ''
    issues.append({'部位': part, '类型': typ, '详情表': dv, '测评表': ev,
                   '差异': diff, '说明': desc, '处理建议': advice, '级别': level,
                   '_group': group, '_tpl': tpl or desc})


# ============================================================
# 单班核验
# ============================================================
def verify_class(cls, cdir):
    """核验一个班，返回 (rows, issues, summary)。对源文件只读。"""
    files = [f for f in glob.glob(os.path.join(cdir, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    detail_path = eval_path = None
    for f in files:
        fn = os.path.basename(f)
        if '详情' in fn:
            detail_path = f
        elif '测评表' in fn:
            eval_path = f

    rows, issues, summary = [], [], None
    if not detail_path or not eval_path:
        issues.append({'班级': cls, '姓名': '', '学号': '', '部位': '文件', '类型': '文件缺失',
                       '详情表': '', '测评表': '', '差异': '', '级别': '必须处理',
                       '说明': f'缺少表格（详情表={os.path.basename(detail_path) if detail_path else "无"}，'
                               f'测评表={os.path.basename(eval_path) if eval_path else "无"}）',
                       '处理建议': '补齐该班两份表后重跑'})
        return rows, issues, {'班级': cls, '学生数': 0, '通过': 0, '不通过': 0,
                              '问题条数': 1, '仅详情': 0, '仅测评': 0}

    last_col_total = detail_last_col_is_total(detail_path)
    if last_col_total:
        print(f"  [提示] {cls} 详情表末列仍为「总分」格式，已按 末列-(思想+文体+科技) 还原志愿服务分")

    try:
        d_roles, d_stu, d_order = load_file(detail_path)
        e_roles, e_stu, e_order = load_file(eval_path)
    except PermissionError as ex:
        issues.append({'班级': cls, '姓名': '', '学号': '', '部位': '文件', '类型': '文件被占用',
                       '详情表': '', '测评表': '', '差异': '', '级别': '必须处理',
                       '说明': f'无法读取（可能被 Excel 打开锁定）：{ex}',
                       '处理建议': '关闭 Excel 后重跑'})
        return rows, issues, {'班级': cls, '学生数': 0, '通过': 0, '不通过': 0,
                              '问题条数': 1, '仅详情': 0, '仅测评': 0}

    d_names, e_names = set(d_stu), set(e_stu)
    matched = d_names & e_names
    only_d, only_e = d_names - e_names, e_names - d_names

    e_by_xh = {norm_xh(e_stu[n]['xuehao']): n for n in e_names}
    d_by_xh = {norm_xh(d_stu[n]['xuehao']): n for n in d_names}
    for n in list(only_d):
        xh = norm_xh(d_stu[n]['xuehao'])
        if xh in e_by_xh and e_by_xh[xh] in only_e:
            matched.add(n)
            matched.add(e_by_xh[xh])
            only_d.discard(n)
            only_e.discard(e_by_xh[xh])
    for n in only_d:
        issues.append({'班级': cls, '姓名': n, '学号': d_stu[n]['xuehao'], '部位': '名单', '类型': '人员缺失',
                       '详情表': '有', '测评表': '无', '差异': '', '级别': '必须处理',
                       '说明': '该生只在详情表中存在，测评表缺此人', '处理建议': '核对是否漏填测评表'})
    for n in only_e:
        issues.append({'班级': cls, '姓名': n, '学号': e_stu[n]['xuehao'], '部位': '名单', '类型': '人员缺失',
                       '详情表': '无', '测评表': '有', '差异': '', '级别': '必须处理',
                       '说明': '该生只在测评表中存在，详情表缺此人', '处理建议': '核对是否漏填详情表'})

    n_match = n_bad = 0
    for n in sorted(matched):
        d, e = d_stu[n], e_stu[n]
        iss = []

        # ---------- 取值 ----------
        sx_e = num(e.get('sixiang_heji'))
        sx_d_sub = sum(num(d.get(k)) for k in ('sixiang_zhengzhi', 'sixiang_pinde', 'sixiang_jilv'))
        sx_qita = sum_desc(d.get('sixiang_qita'))
        if sx_qita is None:
            sx_qita = num(d.get('sixiang_qita'))
        sx_d_sub += sx_qita
        sx_d_raw = num(d.get('sixiang_heji'))
        if sx_d_raw in (0.0, None) and sx_d_sub > 0:
            sx_d = sx_d_sub
            add(iss, '思想品德', '详情表漏填', '', sx_e,
                f'详情表思想品德合计未填（显示0），暂按子项之和 {sx_d_sub} 计',
                '补填详情表思想品德合计列', '必须处理',
                group='sixiang_unfilled', tpl='详情表「思想品德合计」整列未填（显示0），已按子项之和补算')
        else:
            sx_d = sx_d_raw

        wt_d = num(d.get('wenti_heji'))
        wt_e = num(e.get('wenti_jiangzuo'))
        kj_d = num(d.get('keji_heji'))
        kj_e = num(e.get('keji_heji'))
        zy_d = num(d.get('zhiyuan_heji'))
        zy_e = num(e.get('zhiyuan_heji'))
        zy_note = ''
        if last_col_total:
            zy_raw = zy_d
            zy_d = zy_raw - (sx_d + wt_d + kj_d)
            if zy_d < -0.01:
                zy_note = (f'详情表末列(总分列)与前三合计不平：末列={zy_raw}，'
                           f'思想+文体+科技={round(sx_d + wt_d + kj_d, 2)}，还原志愿={round(zy_d, 2)}')
                zy_d = 0.0
            else:
                zy_d = round(zy_d, 2)
        tineng_e = num(e.get('wenti_tineng'))
        e_total = num(e.get('total'))

        # ---------- A. 两表一致性（核心） ----------
        if not approx(sx_d, sx_e):
            add(iss, '思想品德', '两表不一致', sx_d, sx_e,
                f'思想品德合计不符：详情表 {sx_d} vs 测评表 {sx_e}',
                '核对详情表"其他"栏加分与测评表是否一致', '必须处理')
        if not approx(wt_d, wt_e):
            add(iss, '文体活动-讲座', '两表不一致', wt_d, wt_e,
                f'文体活动（讲座、活动）不符：详情表 {wt_d} vs 测评表 {wt_e}',
                '以详情表活动明细为准核对测评表', '必须处理')
        if not approx(kj_d, kj_e):
            add(iss, '科技创新', '两表不一致', kj_d, kj_e,
                f'科技创新合计不符：详情表 {kj_d} vs 测评表 {kj_e}',
                '核对专利/论文/证书等项是否漏填', '必须处理')

        # 志愿服务两表一致性：仅比对两表合计（无数据→需确认，不一致→异常）
        # 末列陷阱班：详情表末列是总分，已还原出真实志愿分，用还原值判断是否为空
        zy_d_is_empty = (zy_d <= 0.001) if last_col_total else vol_empty(d.get('zhiyuan_heji'))
        zy_e_is_empty = vol_empty(e.get('zhiyuan_heji'))
        if zy_d_is_empty and zy_e_is_empty:
            add(iss, '志愿服务', '无数据', zy_d, zy_e,
                '志愿服务两表均无数据（合计为空/无/0），请确认是否属实',
                '确认该生确未参与志愿服务；若实际参与请补填两表志愿合计', '需确认',
                group='zhiyuan_none', tpl='志愿服务两表均无数据，请确认是否属实（未参与志愿服务）')
        elif not approx(zy_d, zy_e):
            add(iss, '志愿服务', '两表不一致', zy_d, zy_e,
                f'志愿服务合计不符：详情表 {zy_d} vs 测评表 {zy_e}',
                '核对两表志愿服务加分是否一致', '必须处理')

        detail_total = round(sx_d + wt_d + kj_d + zy_d, 2)
        total_diff = round(e_total - detail_total, 2)
        zy_gap = round(zy_d - zy_e, 2)
        core_ok = approx(sx_d, sx_e) and approx(wt_d, wt_e) and approx(kj_d, kj_e)
        if approx(total_diff, tineng_e):
            total_ok = True
        elif (not approx(zy_gap, 0.0)) and approx(total_diff, tineng_e - zy_gap) and core_ok:
            total_ok = True     # 差异全部来自志愿服务（已由「志愿服务两表不一致」单独报，避免总分重复提醒）
        else:
            total_ok = False
            add(iss, '总分', '总分不对等', detail_total, e_total,
                f'总分关系不符：测评表总分 {e_total} − 详情表合计 {detail_total} = {total_diff}，'
                f'应等于体能测试 {tineng_e}',
                '修正分项后总分自动对齐', '必须处理')

        # ---------- B. 上限校验（细则） ----------
        for part, val, cap in (('文体活动-讲座', max(wt_d, wt_e), CAP['讲座活动']),
                               ('科技创新-竞赛', num(e.get('keji_jingsai')), CAP['科技竞赛']),
                               ('科技创新', max(kj_d, kj_e), CAP['科技创新合计']),
                               ('思想品德-其他', max(sx_qita, num(e.get('sixiang_qita'))), CAP['其他']),
                               ('思想品德', max(sx_d, sx_e), CAP['思想品德']),
                               ('文体活动', num(e.get('wenti_heji')), CAP['文体活动合计']),
                               ('总计', e_total, CAP['总计'])):
            if val > cap + 0.001:
                add(iss, part, '超出上限', val, '', f'{part}得分 {val} 超过细则上限 {cap} 分',
                    f'按细则封顶为 {cap} 分', '必须处理')
        # ---------- C. 扣分制项（无材料应满分） ----------
        for part, full in DEDUCT_ITEMS.items():
            key = {'政治素质': 'sixiang_zhengzhi', '道德品质': 'sixiang_pinde',
                   '组织纪律': 'sixiang_jilv', '体能测试': 'wenti_tineng'}[part]
            v = num(e.get(key))
            if v < full - 0.001:
                add(iss, part, '扣分需确认', v, '',
                    f'{part}得 {v} 分，低于满分 {full} 分（细则：无扣分材料即记满分）',
                    f'确认该生材料文件夹中有无对应扣分依据', '需确认',
                    group=f'{part}_deduct', tpl=f'「{part}」低于满分 {full} 分（细则：无扣分材料应记满分）')

        # ---------- D. 表内自洽 ----------
        e_sx_sum = sum(num(e.get(k)) for k in ('sixiang_zhengzhi', 'sixiang_pinde', 'sixiang_jilv', 'sixiang_qita'))
        if not approx(sx_e, e_sx_sum):
            add(iss, '思想品德', '测评表不平', '', sx_e,
                f'测评表思想品德合计 {sx_e} ≠ 四个子项之和 {round(e_sx_sum, 2)}',
                '核对测评表合计公式', '必须处理')
        e_wt_heji = num(e.get('wenti_heji'))
        if not approx(e_wt_heji, tineng_e + wt_e):
            add(iss, '文体活动', '测评表不平', '', e_wt_heji,
                f'测评表文体活动合计 {e_wt_heji} ≠ 体能 {tineng_e} + 讲座 {wt_e}',
                '核对测评表合计公式', '必须处理')
        e_kj_sum = sum(num(e.get(k)) for k in ('keji_jingsai', 'keji_lunwen', 'keji_zhuanli',
                                               'keji_wenxue', 'keji_xinwen', 'keji_zige'))
        if not approx(kj_e, e_kj_sum):
            add(iss, '科技创新', '测评表不平', '', kj_e,
                f'测评表科技创新合计 {kj_e} ≠ 六个子项之和 {round(e_kj_sum, 2)}',
                '核对测评表合计公式', '必须处理')
        e_total_sum = round(sx_e + e_wt_heji + kj_e + zy_e, 2)
        if not approx(e_total, e_total_sum):
            add(iss, '总计', '测评表不平', '', e_total,
                f'测评表总计 {e_total} ≠ 各栏合计之和 {e_total_sum}',
                '核对测评表总分公式', '必须处理')
        if not approx(sx_d, sx_d_sub):
            add(iss, '思想品德', '详情表不平', sx_d, '',
                f'详情表思想品德合计 {sx_d} ≠ 子项之和 {round(sx_d_sub, 2)}',
                '核对详情表合计', '必须处理')

        # ---------- 活动明细反算（已停用） ----------
        # 说明：文体活动 / 志愿服务 的「详情表内容反算 ↔ 合计」比较已移除。
        #   用户确认：文体活动明细常混入志愿服务/思政类活动且封顶10分，反算不可靠，
        #   仅当两表合计不一致（见 A 段）才判为错误；志愿服务只比对两表合计（见 A 段）。
        #   科技创新亦只比对两表合计（见 A 段）。
        # （sum_desc 仅保留用于思想品德「其他」栏的 +分值 提取。）

        # ---------- E. 学号 ----------
        xh_d = str(d.get('xuehao')).strip() if d.get('xuehao') is not None else ''
        xh_e = str(e.get('xuehao')).strip() if e.get('xuehao') is not None else ''
        if norm_xh(xh_d) != norm_xh(xh_e) and norm_xh(xh_d) and norm_xh(xh_e):
            add(iss, '学号', '学号不一致', xh_d, xh_e,
                f'两表学号不一致：详情表 {xh_d} vs 测评表 {xh_e}',
                '核对原始学号并统一', '必须处理')

        ok = core_ok and total_ok
        if ok:
            n_match += 1
        else:
            n_bad += 1

        for i in iss:
            rec = {'班级': cls, '姓名': n, '学号': xh_e or xh_d}
            rec.update(i)
            issues.append(rec)

        rows.append({
            '班级': cls, '姓名': n, '学号': xh_e or xh_d,
            '思想品德_详情': sx_d, '思想品德_测评': sx_e,
            '文体活动_详情': wt_d, '讲座活动_测评': wt_e,
            '科技创新_详情': kj_d, '科技创新_测评': kj_e,
            '志愿服务_详情': zy_d, '志愿服务_测评': zy_e,
            '体能测试': tineng_e,
            '详情合计': detail_total, '测评表总分': round(e_total, 2),
            '总分差': total_diff,
            '结果': '通过' if ok else '不通过',
            '问题数': len(iss),
            '问题摘要': '；'.join(f"[{i['部位']}]{i['说明']}" for i in iss) if iss else '',
        })

    summary = {'班级': cls, '学生数': len(matched), '通过': n_match, '不通过': n_bad,
               '问题条数': len(issues), '仅详情': len(only_d), '仅测评': len(only_e)}
    print(f"{cls}: 匹配{len(matched)}人, 通过{n_match}, 不通过{n_bad}")
    return rows, issues, summary


# ============================================================
# 合并整班共性问题（同班同 group 达 5 人以上并成一行，避免刷屏）
# ============================================================
def merge_common_issues(issues_all):
    groups = defaultdict(list)
    for r in issues_all:
        if r.get('_group'):
            groups[(r['班级'], r['_group'])].append(r)
    drop = set()
    merged = []
    for (cls, grp), lst in groups.items():
        if len(lst) < 5:
            continue
        for r in lst:
            drop.add(id(r))
        base = {k: v for k, v in lst[0].items() if not k.startswith('_')}
        names = [x['姓名'] for x in lst]
        show = '、'.join(names[:25]) + (' 等' if len(names) > 25 else '')
        base['姓名'] = f'共 {len(lst)} 人'
        base['学号'] = ''
        base['详情表'] = ''
        base['测评表'] = ''
        base['差异'] = ''
        base['说明'] = lst[0]['_tpl'] + f'　涉及：{show}'
        merged.append(base)
    return [r for r in issues_all if id(r) not in drop] + merged


# ============================================================
# 生成报告（唯一写操作：新建报告簿并保存）
# ============================================================
HDR_FILL = PatternFill('solid', fgColor='305496')
HDR_FONT = Font(bold=True, color='FFFFFF', size=11)
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LEVEL_FILL = {'必须处理': PatternFill('solid', fgColor='FFC7CE'),
              '需确认': PatternFill('solid', fgColor='FFEB9C'),
              '提示': PatternFill('solid', fgColor='EDEDED')}
LEVEL_FONT = {'必须处理': Font(color='9C0006', bold=True),
              '需确认': Font(color='9C6500', bold=True),
              '提示': Font(color='808080')}


def _write_sheet(ws, cols, rows, widths=None, freeze='A2'):
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(1, c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for r in rows:
        ws.append([r.get(c, '') for c in cols])
    for ri in range(1, ws.max_row + 1):
        for ci in range(1, len(cols) + 1):
            ws.cell(ri, ci).border = BORDER
            if ri > 1:
                ws.cell(ri, ci).alignment = Alignment(vertical='center',
                                                      wrap_text=(cols[ci - 1] in ('说明', '处理建议', '问题摘要')))
    if widths:
        for i, cn in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(cn, 12)
    ws.freeze_panes = freeze
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 24


def build_report(out, all_rows, issues_all, summary_classes):
    wb = openpyxl.Workbook()   # 新建报告簿（源表从不被保存）

    # Sheet1 问题清单（主角）
    lv_order = {'必须处理': 0, '需确认': 1, '提示': 2}
    issues_sorted = sorted(issues_all,
                           key=lambda x: (lv_order.get(x.get('级别', '提示'), 3),
                                          x.get('班级', ''), x.get('姓名', '')))
    for i, r in enumerate(issues_sorted, 1):
        r['序号'] = i
    ws1 = wb.active
    ws1.title = '问题清单'
    cols1 = ['序号', '级别', '班级', '姓名', '学号', '部位', '类型', '详情表', '测评表', '差异', '说明', '处理建议']
    _write_sheet(ws1, cols1, issues_sorted,
                 widths={'序号': 5, '级别': 9, '班级': 14, '姓名': 10, '学号': 14, '部位': 13, '类型': 13,
                         '详情表': 9, '测评表': 9, '差异': 8, '说明': 58, '处理建议': 34})
    lv_col = cols1.index('级别') + 1
    for ri in range(2, ws1.max_row + 1):
        v = ws1.cell(ri, lv_col).value
        if v in LEVEL_FILL:
            ws1.cell(ri, lv_col).fill = LEVEL_FILL[v]
            ws1.cell(ri, lv_col).font = LEVEL_FONT[v]
            ws1.cell(ri, lv_col).alignment = Alignment(horizontal='center', vertical='center')

    # Sheet2 全员明细
    ws2 = wb.create_sheet('全员明细')
    cols2 = ['班级', '姓名', '学号', '思想品德_详情', '思想品德_测评', '文体活动_详情', '讲座活动_测评',
             '科技创新_详情', '科技创新_测评', '志愿服务_详情', '志愿服务_测评', '体能测试',
             '详情合计', '测评表总分', '总分差', '结果', '问题数', '问题摘要']
    _write_sheet(ws2, cols2, all_rows,
                 widths={'班级': 14, '姓名': 10, '学号': 14, '总分差': 8, '结果': 8, '问题数': 7, '问题摘要': 70})
    res_col = cols2.index('结果') + 1
    for ri in range(2, ws2.max_row + 1):
        v = ws2.cell(ri, res_col).value
        if v == '不通过':
            ws2.cell(ri, res_col).fill = PatternFill('solid', fgColor='FFC7CE')
            ws2.cell(ri, res_col).font = Font(color='9C0006', bold=True)
        else:
            ws2.cell(ri, res_col).fill = PatternFill('solid', fgColor='C6EFCE')
            ws2.cell(ri, res_col).font = Font(color='006100')
        ws2.cell(ri, res_col).alignment = Alignment(horizontal='center', vertical='center')

    # Sheet3 班级汇总
    ws3 = wb.create_sheet('班级汇总')
    cols3 = ['班级', '学生数', '通过', '不通过', '问题条数', '仅详情', '仅测评']
    _write_sheet(ws3, cols3, summary_classes,
                 widths={'班级': 16, '学生数': 9, '通过': 9, '不通过': 9, '问题条数': 10, '仅详情': 9, '仅测评': 9},
                 freeze='A2')
    ws3.append(['合计', sum(x['学生数'] for x in summary_classes), sum(x['通过'] for x in summary_classes),
                sum(x['不通过'] for x in summary_classes), len(issues_all),
                sum(x['仅详情'] for x in summary_classes), sum(x['仅测评'] for x in summary_classes)])
    for c in range(1, len(cols3) + 1):
        ws3.cell(ws3.max_row, c).font = Font(bold=True)
        ws3.cell(ws3.max_row, c).fill = PatternFill('solid', fgColor='DDEBF7')
    for ri in range(2, ws3.max_row + 1):
        v = ws3.cell(ri, 4).value
        if isinstance(v, int) and v > 0:
            ws3.cell(ri, 4).fill = PatternFill('solid', fgColor='FFC7CE')
            ws3.cell(ri, 4).font = Font(color='9C0006', bold=True)

    # Sheet4 评分规则速查
    ws4 = wb.create_sheet('评分规则速查')
    rules = [
        ['思想品德', '45', '政治素质', '15', '扣分制，无扣分材料即满分'],
        ['思想品德', '45', '道德品质', '10', '扣分制，无扣分材料即满分'],
        ['思想品德', '45', '组织纪律', '10', '扣分制，无扣分材料即满分'],
        ['思想品德', '45', '其他', '10', '加分制：干部(优2.6-3/良2-2.5/合格1-2/宿舍长1)多职取最高；'
                                     '表彰(国家8/省5/校3/院2.5/单项2)多项取最高；献血1分/次≤2；累计≤10'],
        ['文体活动', '20', '体能测试', '10', '扣分制，体育合格即满分10'],
        ['文体活动', '20', '讲座、活动', '10', '加分制：参加活动0.3/次，比赛0.5/次，获奖按级别；上限10'],
        ['科技创新', '25', '科技竞赛活动', '10', '参加校级0.5/次、校级以上1/次；获奖按级别；上限10'],
        ['科技创新', '25', '发表论文', '—', 'SCI/CSSCI 10(8)、EI 8(6)、核心 6(4)、正式刊物 4(2)'],
        ['科技创新', '25', '专利', '—', '发明 授予10/受理5；实用新型 5/3；外观设计 3/1.5；软件著作 4/2'],
        ['科技创新', '25', '文学作品/新闻', '—', '国家级9(7)、省级6(4)、市级3(1)；校院级合计≤1'],
        ['科技创新', '25', '资格证书', '—', '四级2/六级4；计算机二级2/三级4；其他证书2-4；取得起半年内有效'],
        ['志愿服务与社会实践', '10', '志愿服务', '5', '按志愿时长对照表计分，上限5；参与两表一致性审核'],
        ['志愿服务与社会实践', '10', '社会实践', '5', '证明0.5/份、记录表0.2/份、报告1/份；上限5；参与两表一致性审核'],
        ['总分', '100', '—', '—', '详情表文体活动栏不含体能测试(10分)，故测评表总分 − 详情表合计 = 10'],
    ]
    _write_sheet(ws4, ['大栏', '满分', '子项', '上限', '计分要点'],
                 [{'大栏': a, '满分': b, '子项': c, '上限': d, '计分要点': e} for a, b, c, d, e in rules],
                 widths={'大栏': 20, '满分': 8, '子项': 16, '上限': 8, '计分要点': 80}, freeze='A2')

    wb.save(out)   # 仅此处写文件，且写的是新建的报告簿


# ============================================================
# 主流程
# ============================================================
def run_verification(root):
    classes = sorted(d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d)))
    all_rows, issues_all, summary_classes = [], [], []
    for cls in classes:
        rows, issues, summary = verify_class(cls, os.path.join(root, cls))
        all_rows.extend(rows)
        issues_all.extend(issues)
        if summary:
            summary_classes.append(summary)
    # 合并整班共性问题并重算各班问题条数
    issues_all = merge_common_issues(issues_all)
    cnt = Counter(r['班级'] for r in issues_all)
    for s in summary_classes:
        s['问题条数'] = cnt.get(s['班级'], 0)
    return all_rows, issues_all, summary_classes


def main():
    p = argparse.ArgumentParser(description='综测：详情表 vs 测评表 一致性核验')
    p.add_argument('--data', default=DEFAULT_DATA,
                   help='班级文件夹根目录：其下每个子文件夹代表一个班，内含两份 xlsx')
    p.add_argument('--term', default='',
                   help='学期标识，如 2025-2026-2；留空则自动取 data 的上一级目录名')
    p.add_argument('--out', default='',
                   help='报告输出完整路径；留空则自动归档到 <data上级>/核验报告/ 下并按时间命名')
    args = p.parse_args()

    root = args.data
    if not root or not os.path.isdir(root):
        print("用法：python verify.py --data <年级文件夹路径>")
        print("  （GUI 模式下直接运行 run.py，无需此参数）")
        sys.exit(1)
    all_rows, issues_all, summary_classes = run_verification(root)

    # 输出路径：每次只生成一份报告，归档到 <data 上级>/核验报告/<term>_时间戳.xlsx
    term = args.term or os.path.basename(os.path.dirname(os.path.abspath(root)))
    report_dir = os.path.join(os.path.dirname(os.path.abspath(root)), '核验报告')
    os.makedirs(report_dir, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d-%H%M')
    out = args.out or os.path.join(report_dir, f'综测分数核验报告_{term}_{stamp}.xlsx')

    build_report(out, all_rows, issues_all, summary_classes)

    # 控制台汇总
    total = len(all_rows)
    fail = sum(1 for x in all_rows if x.get('结果') == '不通过')
    lv = Counter(x.get('级别') for x in issues_all)
    print('\n' + '=' * 56)
    print(f'核验完成：{total} 人，通过 {total - fail}，不通过 {fail}')
    print(f'问题清单：必须处理 {lv.get("必须处理", 0)} / 需确认 {lv.get("需确认", 0)} / 提示 {lv.get("提示", 0)}')
    print('报告：' + out)
    print('=' * 56)
    print(f"学生 {total} 人；问题 {len(issues_all)} 条  → "
          f"必须处理 {lv.get('必须处理', 0)} / 需确认 {lv.get('需确认', 0)} / 提示(可忽略) {lv.get('提示', 0)}")

    # 机器可读汇总（供 run.py 子进程模式解析；不影响人读输出）
    import json
    print("SUMMARY_JSON:" + json.dumps({"total": total, "fail": fail, "lv": dict(lv)}, ensure_ascii=False))


if __name__ == '__main__':
    main()
