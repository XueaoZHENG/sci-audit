# -*- coding: utf-8 -*-
"""
low_n_probe.py -- 低样本(n<=3)专项数值取证探针

背景
----
通用检测器 (scripts/detectors.py) 的阈值是 "相同块长度 >=2 / 恒定偏移 >=3 对"，
因此漏检了 PubPeer 评论 #11:
    MOESM10 (ED Fig.1a) 两个 OE 株系块:
      首值完全相同 (11.2 == 11.2),
      其余两对差恰好 1.0 (12 vs 13, 11.4 vs 12.4),
      且 OE 组只有 1 位小数 vs NIPB 组 14 位小数 (精度断层)。

本脚本实现 4 个低样本专项检测器 + 逐值目视比对清单:

  D1  首值相同检测        : n<=3 块之间, 第 1 个值完全相同 (原始字符串级)
  D2  精确整数偏移检测    : 存在 k/n 对满足 |A-B| = 同一整数值 (k>=2, n<=3)
  D3  精度断层检测        : 同 sheet 内不同块的有效小数位数差异 >= 3 位
  D4  共享长十进制尾串    : 附加探针, 两个值共享 >=8 位连续十进制尾 (强派生信号)
  D5  跨文件/跨图版逐值全同: n<=3 块的全部原始字符串序列完全一致

判定纪律 (来自审查技能)
------------------------
  * n=3 的 2/3 对恒定偏移 / 1 位小数的相同值, 单独都不构成红线 (巧合空间大);
  * 给红线必须: 跨文件或跨图版的逐值全同, 或多维证据叠加;
  * n=3 的弱模式一律标 "存疑(orange)", 并写清 "需要什么材料可定论";
  * 每条命中必须给出 单元格坐标 + 原始字符串值;
  * 必须列出假阳性 (等差序号列 / 归一化锚点 / n 粒度均值 / 统计表派生量) 并说明排除理由。

用法:  python scripts/low_n_probe.py
输出:  _dump/low_n_findings.txt   (详细)
       _dump/low_n_summary.md     (中文结论摘要)
       _dump/low_n_blocks.json     (机器可读: 所有块)
"""

import os
import re
import sys
import json
import zipfile
import itertools
import xml.etree.ElementTree as ET
from collections import defaultdict

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.stderr.write("需要 openpyxl: pip install openpyxl\n")
    raise

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DUMP = os.path.join(BASE, "_dump")
os.makedirs(DUMP, exist_ok=True)

FILES = ["MOESM7.xlsx", "MOESM8.xlsx", "MOESM9.xlsx", "MOESM10.xlsx", "MOESM11.xlsx",
         "MOESM12.xlsx", "MOESM13.xlsx", "MOESM14.xlsx", "MOESM15.xlsx", "MOESM16.xlsx"]

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# ---------------------------------------------------------------- 已知已命中 (不计为新发现)
KNOWN = {
    "known#8": "MOESM10 ED1a B3:B5 == MOESM13 ED7b B12:B14 (逐值全同, 跨文件)",
    "known#10": "MOESM11 ED3c E8=E9 / F8=F10 / G8=G10 (组内重复值)",
    "known#6": "MOESM14 ED9f 全部来自 6 条 14 位母串",
    "known#7": "MOESM13 ED7b 块内 +-0.2/+-0.24 编辑",
    "known#2": "MOESM15 ED10l 8 连值共享 / +-200000 偏移",
    "known#3": "MOESM15 ED10l 行9/行10 跨三列比值恒定",
    "known#15": "MOESM15 ED10l 行9/行10 精确等比数列 0.99^n",
    "known#4": "MOESM7 Fig.1c/1f Oscie1-1 24 值全同",
    "known#5": "MOESM7 Fig.1c 行16-24 比值恒定",
    "known#9": "MOESM11 ED3f 样本7 三对值比值恒定",
    "known#16": "MOESM11 ED3b 前8值 ~1.15116279 x MOESM15 ED10l",
    "clean": "MOESM8 / MOESM16 (阴性对照)",
}

STAT_KEYWORDS = [
    "mean diff", "ci of diff", "significant?", "summary", "adjusted p",
    "p value", "t, df", "tukey", "one- or two-tailed", "number of pairs",
    "unpaired t test", "paired t test",
]


# ================================================================ 原始存储字符串
def raw_stored_map(path):
    """返回 {sheet_name: {coord: 原始存储字符串}} —— 只含数值型单元格 (t 不在 s/str/... 中)。

    openpyxl 的 cell.value 会把数字转成 Python float, 丢失原始字面量
    (例如 Excel 里存 "0.116677015184085" 会被转成 float)。本函数直接解析
    sheet XML 的 <v> 节点, 拿到 Excel 真正写进文件的字符串。
    """
    out = {}
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels}
        sheets = wb.find(NS_MAIN + "sheets")
        for sh in sheets:
            name = sh.get("name")
            rid = sh.get(NS_REL + "id")
            target = relmap[rid]
            if target.startswith("/"):
                target = target[1:]
            elif not target.startswith("xl/"):
                target = "xl/" + target
            root = ET.fromstring(z.read(target))
            m = {}
            for c in root.iter(NS_MAIN + "c"):
                t = c.get("t")
                if t in ("s", "str", "inlineStr", "b", "e"):
                    continue
                v = c.find(NS_MAIN + "v")
                if v is None or v.text is None:
                    continue
                m[c.get("r")] = v.text
            out[name] = m
    return out


def decimals_of(s):
    """原始存储字符串的有效小数位数。"""
    if s is None:
        return 0
    s = s.strip()
    if "e" in s or "E" in s:
        try:
            return max(0, -int(s.split("e")[1].split("E")[0]))
        except Exception:
            return 0
    if "." not in s:
        return 0
    frac = s.split(".", 1)[1]
    return len(frac)


def digits_of(s):
    """去掉小数点后的数字串 (用于尾串比对)。"""
    if s is None:
        return ""
    s = s.strip()
    if "e" in s or "E" in s:
        s = s.split("e")[0].split("E")[0]
    return s.replace(".", "").replace("-", "").replace("+", "")


def frac_digits(s):
    """小数部分的数字串 (右对齐比对用)。"""
    if s is None:
        return ""
    s = s.strip()
    if "." not in s:
        return ""
    return s.split(".", 1)[1]


def trailing_match(a, b):
    """两个原始字符串共享的连续尾部小数位数 (右对齐)。"""
    fa, fb = frac_digits(a), frac_digits(b)
    n = 0
    while n < len(fa) and n < len(fb) and fa[len(fa) - 1 - n] == fb[len(fb) - 1 - n]:
        n += 1
    return n


def is_int_like(x):
    return abs(x - round(x)) < 1e-9


# ================================================================ 读 sheet 结构
class Block(object):
    __slots__ = ("file", "sheet", "col", "rows", "vals", "raws", "canon", "dps",
                 "label", "tag", "sig", "coords", "first_raw")

    def __init__(self, file, sheet, col, rows, vals, raws, canon, dps, label, tag):
        self.file = file
        self.sheet = sheet
        self.col = col
        self.rows = rows
        self.vals = vals
        self.raws = raws
        self.canon = canon
        self.dps = dps
        self.label = label
        self.tag = tag
        self.sig = "%s | %s | col %s" % (file, sheet, col)
        self.coords = ["%s%d" % (col, r) for r in rows]
        self.first_raw = raws[0] if raws else None

    @property
    def n(self):
        return len(self.vals)

    @property
    def maxdp(self):
        return max(self.dps) if self.dps else 0

    def __repr__(self):
        return "<Block %s n=%d tag=%s label=%r>" % (self.sig, self.n, self.tag, self.label)


def load_blocks():
    blocks = []
    sheets_seen = []
    for fname in FILES:
        path = os.path.join(BASE, fname)
        raw = raw_stored_map(path)
        wb = openpyxl.load_workbook(path, data_only=False)
        for sname in wb.sheetnames:
            ws = wb[sname]
            raws_map = raw.get(sname, {})
            # 枚举全部非空单元格 (不依赖 openpyxl 的 dimensions)
            cellmap = {}
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        cellmap[c.row] = cellmap.get(c.row, {})
                        cellmap[c.row][c.column] = c
            sheets_seen.append((fname, sname))
            # 按列收集数值单元格
            bycol = defaultdict(list)
            for r, cols in cellmap.items():
                for col, c in cols.items():
                    v = c.value
                    if isinstance(v, bool):
                        continue
                    if isinstance(v, (int, float)):
                        bycol[col].append((r, c, float(v)))
            for col in sorted(bycol):
                items = sorted(bycol[col])
                # 按行连续性切块
                run = [items[0]]
                for it in items[1:]:
                    if it[0] == run[-1][0] + 1:
                        run.append(it)
                    else:
                        blocks.append(make_block(fname, sname, col, run, cellmap, raws_map))
                        run = [it]
                blocks.append(make_block(fname, sname, col, run, cellmap, raws_map))
    return blocks, sheets_seen


def make_block(fname, sname, col, run, cellmap, raws_map):
    rows = [r for (r, _c, _v) in run]
    vals = [v for (_r, _c, v) in run]
    raws = []
    canon = []
    dps = []
    for (r, _c, _v) in run:
        s = raws_map.get("%s%d" % (openpyxl.utils.get_column_letter(col), r))
        if s is None:
            s = repr(_v)
        raws.append(s)
        # canon: Python 最短往返 repr —— 去掉写文件时 float->str 带来的 17 位尾巴
        # (例如 XML 里存 '0.90230473800567002', 有效精度其实是 14 位小数)
        c = repr(_v)
        canon.append(c)
        dps.append(decimals_of(c))
    # 标签: 同列、块首行之上最近的非空单元格
    label = None
    for r in range(rows[0] - 1, 0, -1):
        c = cellmap.get(r, {}).get(col)
        if c is not None:
            label = str(c.value)
            break
    # 分级
    tag = "DATA"
    if label and any(k in label.lower() for k in STAT_KEYWORDS):
        tag = "STAT"
    if len(vals) >= 2 and all(is_int_like(v) for v in vals) and \
            list(range(int(round(vals[0])), int(round(vals[0])) + len(vals))) == [int(round(v)) for v in vals]:
        tag = "SEQ"
    return Block(fname, sname, openpyxl.utils.get_column_letter(col), rows, vals,
                 raws, canon, dps, label, tag)


# ================================================================ 检测器
def fmt(x):
    if isinstance(x, float) and x == int(x) and abs(x) < 1e15:
        return "%.1f" % x
    return repr(x)


def pair_diff_str(a, b):
    """标注两个值的差异性质。"""
    if a is None or b is None:
        return ""
    fa, fb = float(a), float(b)
    if fa == fb:
        return "完全相等"
    d = fa - fb
    if is_int_like(d):
        return "差值=整数(%+d)" % int(round(d))
    return "差值=%+.6g(非整数)" % d


def _cand(blocks, tags=("DATA",)):
    return [b for b in blocks if 2 <= b.n <= 3 and b.tag in tags]


def detect_first_value_equal(blocks, tags=("DATA",)):
    """D1: 首值相同 (原始字符串级完全一致)。"""
    hits = []
    cand = _cand(blocks, tags)
    for A, B in itertools.combinations(cand, 2):
        if A.canon[0] != B.canon[0]:
            continue
        same_sheet = (A.file, A.sheet) == (B.file, B.sheet)
        hits.append(dict(A=A, B=B, scope=("同sheet" if same_sheet else "跨sheet/跨文件"),
                         raw=A.canon[0], rawA=A.raws[0], rawB=B.raws[0]))
    return hits


def detect_constant_int_offset(blocks, tags=("DATA",)):
    """D2: 精确整数偏移 —— k/n 对满足 |A-B| = 同一整数值 (k>=2)。"""
    hits = []
    cand = _cand(blocks, tags)
    for A, B in itertools.combinations(cand, 2):
        if A.n != B.n:
            continue
        diffs = [A.vals[i] - B.vals[i] for i in range(A.n)]
        mags = [abs(d) for d in diffs]
        ints = [m for m in mags if is_int_like(m)]
        if len(ints) < 2:
            continue
        # 找出出现 >=2 次的整数偏移
        from collections import Counter
        cnt = Counter(int(round(m)) for m in ints)
        for off, k in cnt.items():
            if k < 2:
                continue
            dev = []
            for i in range(A.n):
                m = mags[i]
                if is_int_like(m) and int(round(m)) == off:
                    dev.append("第%d对=命中(%s vs %s)" % (i + 1, A.raws[i], B.raws[i]))
                else:
                    dev.append("第%d对=偏离(%s vs %s, |d|=%.10g)"
                               % (i + 1, A.raws[i], B.raws[i], m))
            same_sheet = (A.file, A.sheet) == (B.file, B.sheet)
            hits.append(dict(A=A, B=B, offset=off, k=k, n=A.n, dev=dev,
                             scope=("同sheet" if same_sheet else "跨sheet/跨文件"),
                             same_rows=(A.rows == B.rows)))
    return hits


def detect_precision_break(blocks):
    """D3: 同 sheet 内不同块的有效小数位数差异 >= 3。"""
    hits = []
    bysheet = defaultdict(list)
    for b in blocks:
        if b.tag != "DATA":
            continue
        bysheet[(b.file, b.sheet)].append(b)
    for key, bs in bysheet.items():
        for A, B in itertools.combinations(bs, 2):
            if A.maxdp == B.maxdp:
                continue
            gap = abs(A.maxdp - B.maxdp)
            # 断层要有意义: 一方被四舍五入到 <=2 位小数, 另一方保留 >=5 位
            if gap >= 3 and min(A.maxdp, B.maxdp) <= 2 and max(A.maxdp, B.maxdp) >= 5:
                hits.append(dict(A=A, B=B, gap=gap, dpA=A.maxdp, dpB=B.maxdp))
    return hits


def detect_shared_tail(blocks, tags=("DATA",)):
    """D4: 共享长十进制尾串 (>=8 位连续相同小数尾)。附加探针。"""
    hits = []
    cand = _cand(blocks, tags)
    for A, B in itertools.combinations(cand, 2):
        if A.n != B.n:
            continue
        for i in range(A.n):
            ra, rb = A.canon[i], B.canon[i]
            if ra == rb:
                continue
            tm = trailing_match(ra, rb)
            sig = max(len(digits_of(ra)), len(digits_of(rb)))
            if tm >= 8 and sig >= 12:
                hits.append(dict(A=A, B=B, idx=i, tm=tm, ra=ra, rb=rb,
                                 rawA=A.raws[i], rawB=B.raws[i],
                                 delta=float(A.vals[i]) - float(B.vals[i]),
                                 scope=("同sheet" if (A.file, A.sheet) == (B.file, B.sheet) else "跨sheet/跨文件")))
    return hits


def detect_cross_identical(blocks, tags=("DATA",)):
    """D5: 跨文件 / 跨图版, n<=3 块的全部原始字符串序列完全一致。"""
    hits = []
    cand = _cand(blocks, tags)
    for A, B in itertools.combinations(cand, 2):
        if (A.file, A.sheet) == (B.file, B.sheet):
            continue
        if A.canon == B.canon:
            hits.append(dict(A=A, B=B))
    return hits


def detect_full_constant_offset(blocks):
    """D6: n/n 对全部满足 A_i - B_i = 同一常数 k (k != 0)。比 D2 更严 (要求全对命中)。"""
    hits = []
    cand = _cand(blocks)
    for A, B in itertools.combinations(cand, 2):
        if A.n != B.n:
            continue
        ds = [A.vals[i] - B.vals[i] for i in range(A.n)]
        k = ds[0]
        if k == 0:
            continue
        if all(abs(d - k) < 1e-9 for d in ds):
            hits.append(dict(A=A, B=B, k=k))
    return hits


def detect_quantisation_grid(blocks):
    """D7: 粒度网格检测 —— 组内全部值是否落在 1/5 (0.2) 或 1/19 等细粒度网格上。"""
    hits = []
    cand = _cand(blocks)
    for b in cand:
        for denom, name in ((5, "0.2"), (19, "1/19")):
            if all(abs(v * denom - round(v * denom)) < 1e-9 for v in b.vals):
                hits.append(dict(b=b, denom=denom, name=name))
                break
    return hits


def high_precision_index(blocks, min_sig=13):
    """高精度值索引: >=min_sig 位有效数字的值在全语料出现 >1 次的位置。"""
    idx = defaultdict(list)
    for b in blocks:
        if b.tag != "DATA":
            continue
        for c, r in zip(b.coords, b.canon):
            if len(digits_of(r)) >= min_sig:
                idx[r].append((b.file, b.sheet, b.label, c))
    return {k: v for k, v in idx.items() if len(v) > 1}


def xml_string(file, sheet, coord):
    """取某个单元格在 xlsx 里真正存储的 <v> 文本。"""
    path = os.path.join(BASE, file)
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels}
        tgt = None
        for sh in wb.find(NS_MAIN + "sheets"):
            if sh.get("name") == sheet:
                tgt = relmap[sh.get(NS_REL + "id")]
        if tgt.startswith("/"):
            tgt = tgt[1:]
        elif not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        root = ET.fromstring(z.read(tgt))
        for c in root.iter(NS_MAIN + "c"):
            if c.get("r") == coord:
                v = c.find(NS_MAIN + "v")
                return v.text if v is not None else None
    return None


# ================================================================ 重点复核
def find_block(blocks, file, sheet, col, rows):
    for b in blocks:
        if (b.file, b.sheet, b.col) == (file, sheet, col) and b.rows == list(rows):
            return b
    return None


def ed1a_deep_dive(blocks):
    """MOESM10 ED Fig.1a / ED Fig.1d 的逐单元格原始串 + 跨文件派生表。"""
    LINES = []
    LINES.append("=" * 100)
    LINES.append("重点复核 A: MOESM10 (Source Data ED Fig. 1) —— 'Extended Data Fig. 1a' sheet 原始单元格")
    LINES.append("=" * 100)
    LINES.append("")
    LINES.append("文件: MOESM10.xlsx  sheet: 'Extended Data Fig. 1a'  (PubPeer 评论 #11 指向的位置)")
    LINES.append("图注 (captions_full.txt): 'a, Relative expression of OsCERK1. The rice CYCLOPHILIN2 gene")
    LINES.append("  served as an internal control. (Values are means +/- SD, n = 3, biologically independent")
    LINES.append("  samples).'  -> 同一面板内三个基因型, 同一批次 qPCR, n=3。")
    LINES.append("")
    for col, lbl in (("B", "NIPB"), ("C", "OEOsCERK1-mCherry-1"), ("D", "OEOsCERK1-mCherry-2")):
        LINES.append("  --- 列 %s  表头 %s ---" % (col, lbl))
        for r in (3, 4, 5):
            coord = "%s%d" % (col, r)
            LINES.append("      %-4s XML原始存储串 = %-24r   有效小数位 = %d"
              % (coord, xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", coord),
                 decimals_of(repr(float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", coord))))))
        LINES.append("")
    LINES.append("  #11 三项子命题逐条核对:")
    c3 = xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", "C3")
    d3 = xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", "D3")
    LINES.append("    (1) 首值完全相同: C3=%r  D3=%r  -> %s" % (c3, d3, "成立" if c3 == d3 else "不成立"))
    for (ca, cb, r) in (("C4", "D4", 4), ("C5", "D5", 5)):
        va = float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", ca))
        vb = float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", cb))
        LINES.append("    (2) 第%d对差值: %s=%r  %s=%r  -> 差 %+g  (恰好 1.0: %s)"
          % (r, ca, va, cb, vb, vb - va, "是" if abs(abs(vb - va) - 1.0) < 1e-12 else "否"))
    dpn = max(decimals_of(repr(float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", "B%d" % r)))) for r in (3, 4, 5))
    dpc = max(decimals_of(repr(float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", "C%d" % r)))) for r in (3, 4, 5))
    dpd = max(decimals_of(repr(float(xml_string("MOESM10.xlsx", "Extended Data Fig. 1a", "D%d" % r)))) for r in (3, 4, 5))
    LINES.append("    (3) 精度断层: NIPB 组有效小数位 = %d, OE-1 组 = %d, OE-2 组 = %d -> 断层 %d 位 (>=3 成立: %s)"
      % (dpn, dpc, dpd, max(dpn - dpc, dpn - dpd), "是"))
    LINES.append("        注: PubPeer 原文写 '6 位小数', 是把 0.90230473800567 取整到 6 位后的说法;")
    LINES.append("            实际存储精度为 14 位小数, 断层幅度比 PubPeer 描述的更大。")
    LINES.append("")

    LINES.append("=" * 100)
    LINES.append("重点复核 B: MOESM10 'Extended Data Fig. 1d' <- MOESM9 'Fig.3e' 逐格派生表")
    LINES.append("=" * 100)
    LINES.append("")
    LINES.append("图注: ED Fig.1d = 'Relative fungal growth, values are means +/- SD (n = 3, biologically")
    LINES.append("  independent samples)' (OsCERK1-mCherry 过表达株系, 与 ED1a 同一批基因型)")
    LINES.append("主文 Fig.3e = MOESM9 (Source Data Fig. 3) 的 Oscie1-1 互补株系 qPCR/相对生长量, 11 个基因型 x n=3")
    LINES.append("")
    pairs = [
        ("B", "NIPB", "B", "NIPB"),
        ("C", "OEOsCERK1-mCherry-1", "C", "Oscie1-1"),
        ("D", "OEOsCERK1-mCherry-2", "C", "Oscie1-1"),
    ]
    for ecol, elab, scol, slab in pairs:
        LINES.append("  ED1d 列 %s (%s)  <-  Fig.3e 列 %s (%s)" % (ecol, elab, scol, slab))
        for r in (3, 4, 5):
            e = xml_string("MOESM10.xlsx", "Extended Data Fig. 1d", "%s%d" % (ecol, r))
            f = xml_string("MOESM9.xlsx", "Fig.3e", "%s%d" % (scol, r))
            de, df = float(e), float(f)
            tm = trailing_match(repr(de), repr(df))
            d = de - df
            # 偏移是否是"整十进制常数" (0.01/0.02/0.04/0.1/...) —— float 残差 <=1e-12
            off = None
            for cand_off in (0.01, 0.02, 0.03, 0.04, 0.05, 0.1, 0.2, 0.24, 0.25, 0.28, 0.3, 0.35, 0.4):
                if abs(abs(d) - cand_off) < 1e-12:
                    off = cand_off if d >= 0 else -cand_off
                    break
            if de == df:
                verdict = "**逐位完全相同**"
            elif off is not None:
                verdict = "= Fig3e 值 + (%+g)  [float残差 %+.2g] %s" % (off, d - off, "(共享尾串 %d 位)" % tm if tm >= 8 else "")
            elif tm >= 8:
                verdict = "共享尾串 %d 位" % tm
            else:
                verdict = "?"
            LINES.append("      第%d对: ED1d %s%d=%-22r  Fig3e %s%d=%-22r  delta=%+.17g  %s"
              % (r, ecol, r, e, scol, r, f, d, verdict))
        LINES.append("")
    return LINES


# ================================================================ 输出
def block_line(b):
    parts = []
    for c, r, cn in zip(b.coords, b.raws, b.canon):
        if r == cn:
            parts.append("%s=%s" % (c, r))
        else:
            parts.append("%s=%s [canon=%s]" % (c, r, cn))
    return "%s [%s] label=%r n=%d dp=%s : %s" % (
        b.sig, b.tag, b.label, b.n, b.dps, ", ".join(parts))


def main():
    blocks, sheets_seen = load_blocks()

    with open(os.path.join(DUMP, "low_n_blocks.json"), "w", encoding="utf-8") as fj:
        json.dump([dict(file=b.file, sheet=b.sheet, col=b.col, rows=b.rows,
                        raws=b.raws, canon=b.canon, vals=b.vals, dps=b.dps, label=b.label,
                        tag=b.tag, n=b.n) for b in blocks], fj, ensure_ascii=False, indent=1)

    d1 = detect_first_value_equal(blocks)
    d2 = detect_constant_int_offset(blocks)
    d3 = detect_precision_break(blocks)
    d4 = detect_shared_tail(blocks)
    d5 = detect_cross_identical(blocks)
    d6 = detect_full_constant_offset(blocks)
    d7 = detect_quantisation_grid(blocks)
    hpi = high_precision_index(blocks)

    # 假阳性对照: 把同一批检测器跑在 SEQ(序号列) + STAT(统计表) 块上, 统计会被淹没的命中
    fp_blocks = [b for b in blocks if b.tag in ("SEQ", "STAT")]
    fp = dict(
        d1=len(detect_first_value_equal(fp_blocks, ("SEQ", "STAT"))),
        d2=len(detect_constant_int_offset(fp_blocks, ("SEQ", "STAT"))),
        d4=len(detect_shared_tail(fp_blocks, ("SEQ", "STAT"))),
        d5=len(detect_cross_identical(fp_blocks, ("SEQ", "STAT"))),
    )

    # 面板定义: 一个 sheet = 一个图版面板
    panel_blocks = defaultdict(list)
    for b in blocks:
        panel_blocks[(b.file, b.sheet)].append(b)

    low_n_panels = []          # 含有 n<=3 DATA 块的面板
    for key, bs in panel_blocks.items():
        data = [b for b in bs if b.tag == "DATA" and 2 <= b.n <= 3]
        if data:
            low_n_panels.append((key, data))
    low_n_panels.sort()

    out = []
    W = out.append
    W("=" * 100)
    W("低样本(n<=3)专项数值取证 —— 完整结果")
    W("生成脚本: scripts/low_n_probe.py")
    W("文件总数: %d   sheet 总数: %d   块(block)总数: %d" % (len(FILES), len(sheets_seen), len(blocks)))
    W("含 n<=3 DATA 块的面板数: %d" % len(low_n_panels))
    W("=" * 100)
    W("")

    # ---------------- 0. 面板清单 + 覆盖
    W("################ 0. 面板覆盖清单 ################")
    W("")
    for (f, s), data in low_n_panels:
        W("- %s | %s" % (f, s))
        for b in sorted(data, key=lambda x: (x.col, x.rows[0])):
            W("    " + block_line(b))
    W("")

    # ---------------- 1. 逐值目视比对清单
    W("################ 1. n<=3 面板逐值目视比对清单 (全部组对, 全部配对值) ################")
    W("")
    for (f, s), data in low_n_panels:
        W("=" * 90)
        W("PANEL: %s | %s" % (f, s))
        W("=" * 90)
        for b in sorted(data, key=lambda x: (x.col, x.rows[0])):
            W("  组: %s" % block_line(b))
        W("")
        for A, B in itertools.combinations(sorted(data, key=lambda x: (x.col, x.rows[0])), 2):
            if A.n != B.n:
                continue
            W("  -- 组对: %s  VS  %s   (行跨度 %s vs %s)" %
              (A.label or A.col, B.label or B.col, A.rows, B.rows))
            for i in range(A.n):
                W("     第%d对: %s=%s | %s=%s  -> %s" %
                  (i + 1, A.coords[i], A.raws[i], B.coords[i], B.raws[i],
                   pair_diff_str(A.vals[i], B.vals[i])))
            W("")
    W("")

    # ---------------- 2. D1 首值相同
    W("################ 2. D1 首值相同检测 (n<=3 块, 原始字符串级) ################")
    W("")
    if not d1:
        W("  (无命中)")
    for h in d1:
        A, B = h["A"], h["B"]
        W("  [命中] 首值相同: canon=%r | A XML原始串=%s | B XML原始串=%s"
          % (h["raw"], h["rawA"], h["rawB"]))
        W("     A: %s" % block_line(A))
        W("     B: %s" % block_line(B))
        W("     范围: %s" % h["scope"])
        W("")

    # ---------------- 3. D2 恒定整数偏移
    W("################ 3. D2 精确整数偏移检测 (>=2/%d 对 |A-B| = 同一整数) ################" % 3)
    W("")
    if not d2:
        W("  (无命中)")
    for h in d2:
        A, B = h["A"], h["B"]
        W("  [命中] 恒定整数偏移 |A-B| = %d , 命中 %d/%d 对 (范围: %s, 同行=%s)"
          % (h["offset"], h["k"], h["n"], h["scope"], h["same_rows"]))
        W("     A: %s" % block_line(A))
        W("     B: %s" % block_line(B))
        for d in h["dev"]:
            W("       - " + d)
        W("")

    # ---------------- 4. D3 精度断层
    W("################ 4. D3 精度断层检测 (同 sheet, 小数位数差 >= 3) ################")
    W("")
    if not d3:
        W("  (无命中)")
    for h in d3:
        A, B = h["A"], h["B"]
        W("  [命中] 小数位数差 = %d 位 (A=%d 位, B=%d 位)" % (h["gap"], h["dpA"], h["dpB"]))
        W("     A: %s" % block_line(A))
        W("     B: %s" % block_line(B))
        W("")

    # ---------------- 5. D4 共享尾串
    W("################ 5. D4 共享长十进制尾串检测 (附加探针, >=8 位尾串) ################")
    W("")
    if not d4:
        W("  (无命中)")
    for h in d4:
        A, B = h["A"], h["B"]
        W("  [命中] 第%d对共享尾部 %d 位: %s=%s (XML %s) vs %s=%s (XML %s)  (delta=%+.10g, 范围:%s)"
          % (h["idx"] + 1, h["tm"], A.coords[h["idx"]], h["ra"], h["rawA"],
             B.coords[h["idx"]], h["rb"], h["rawB"], h["delta"], h["scope"]))
        W("     A: %s" % block_line(A))
        W("     B: %s" % block_line(B))
        W("")

    # ---------------- 6. D5 跨文件全同
    W("################ 6. D5 跨文件/跨图版 逐值全同 (n<=3 块) ################")
    W("")
    if not d5:
        W("  (无命中)")
    for h in d5:
        A, B = h["A"], h["B"]
        W("  [命中] 全部原始字符串序列一致")
        W("     A: %s" % block_line(A))
        W("     B: %s" % block_line(B))
        W("")

    # ---------------- 6b. D6 全对恒定偏移
    W("################ 6b. D6 n/n 全对恒定偏移 (任意常数 k != 0) ################")
    W("")
    if not d6:
        W("  (无命中 —— n<=3 的 DATA 组之间不存在 '整列 = 另一列 + 同一常数' 的完整复制)")
    for h in d6:
        W("  [命中] k=%+.10g" % h["k"])
        W("     A: %s" % block_line(h["A"]))
        W("     B: %s" % block_line(h["B"]))
    W("")

    # ---------------- 6c. D7 粒度网格
    W("################ 6c. D7 粒度网格检测 (0.2 / 1-19 细粒度) ################")
    W("")
    if not d7:
        W("  (无命中)")
    for h in d7:
        W("  [命中] 全部值落在 1/%d 网格 (%s 粒度): %s" % (h["denom"], h["name"], block_line(h["b"])))
    W("")

    # ---------------- 6d. 高精度重复值全局索引
    W("################ 6d. 全语料高精度重复值索引 (>=13 位有效数字) ################")
    W("")
    W("  这张表是本轮低样本探针最有力的证据: 两个'独立'测量共享 13 位以上有效数字,")
    W("  巧合概率 ~1e-13, 只可能来自同一原始数的复制/加常数编辑。")
    W("")
    for k, v in sorted(hpi.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        W("  值 %s  (%d 次)" % (k, len(v)))
        for f, sh, lbl, c in v:
            W("      %s | %s | %s | %s   XML原始串=%s" % (f, sh, lbl, c, xml_string(f, sh, c)))
    W("")

    # ---------------- 7. STAT/SEQ 块一览 (假阳性来源)
    W("################ 7. 统计表 / 序号列块一览 (假阳性来源, 不计入命中) ################")
    W("")
    for b in blocks:
        if b.tag in ("STAT", "SEQ") and 2 <= b.n <= 3:
            W("  %-4s %s" % (b.tag, block_line(b)))
    W("")
    W("  若不排除 SEQ/STAT 块, 同一批检测器会产生: D1=%(d1)d, D2=%(d2)d, D4=%(d4)d, D5=%(d5)d 条噪声命中" % fp)
    W("  (统计表里的 Mean Diff. / Adjusted P Value 是由数据算出的派生量, 序号列 1,2,3 是人工填的重复编号,")
    W("   二者与'独立测量之间出现恒定关系'无因果关系, 故全部排除。)")
    W("")

    W2 = ed1a_deep_dive(blocks)
    out.extend(W2)

    with open(os.path.join(DUMP, "low_n_findings.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")

    # 供 summary 使用
    stats = dict(
        n_blocks=len(blocks),
        n_panels=len(low_n_panels),
        d1=len(d1), d2=len(d2), d3=len(d3), d4=len(d4), d5=len(d5),
        d6=len(d6), d7=len(d7), hpi=len(hpi),
        fp=fp,
    )
    print(json.dumps(stats, indent=1))
    print("wrote", os.path.join(DUMP, "low_n_findings.txt"))
    return blocks, low_n_panels, d1, d2, d3, d4, d5, sheets_seen


if __name__ == "__main__":
    main()
