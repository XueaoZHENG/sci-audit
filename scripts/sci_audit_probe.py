#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sci-audit 数值/元数据探针（v2 合并版）
================================================================
一条命令跑完 T0（材料+元数据）与 T1/T2（数值指纹 + 低样本 + 统计自洽），
所有输出落文件、不进上下文 —— 这是本技能控制 token 成本的核心。

用法:
    python sci_audit_probe.py --input <含 xlsx 的目录> --out <输出目录>
                              [--control <已确认干净的面板名关键词>]

产出:
    dump/cells.jsonl          全量非空单元格（机读）
    dump/<file>_dump.txt      每文件逐格文本转储
    findings_numeric.txt      T1 数值指纹探针（P1–P8）
    findings_lown.txt         T2 低样本 n<=3 专项（L1–L5）
    findings_meta.txt         T0 容器与元数据取证
    findings_stats.txt        T2 统计量自洽性复算（Mean Diff. / t）
    summary.md                给人看的摘要（先读这个）

依赖: numpy, scipy, openpyxl
"""
import argparse, json, os, re, sys, zipfile, math, datetime
from collections import defaultdict
from fractions import Fraction

import numpy as np
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

# ----------------------------------------------------------------------
# 0. 原始值提取：直接解析 xl/worksheets/sheetN.xml 取 <v> 字面量
#    （openpyxl 的 cell.value 只给 float，会丢掉原始串与有效位数）
# ----------------------------------------------------------------------
# 注意：必须能区分自闭合的 <c r="A2" s="2"/> 与 <c r="B2" ...>...</c>。
# 若第一个分支允许 attrs 吞掉结尾的 "/"，自闭合标签会吃掉下一格的内容，造成静默丢格。
RE_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
RE_V = re.compile(r"<v>([^<]*)</v>")
RE_F = re.compile(r"<f[^>]*>(.*?)</f>", re.S)
RE_T = re.compile(r'<t[^>]*>(.*?)</t>', re.S)
RE_SST = re.compile(r"<si>(.*?)</si>", re.S)
RE_SST_T = re.compile(r"<t[^>]*>(.*?)</t>", re.S)

# GraphPad 统计块表头关键词 —— 用于把块分类为 STAT 并剔除，避免海量假阳性
STAT_HEADER_KW = ("mean diff", "adjusted p", "p value", "p value summary",
                  "95.00% ci", "significant?", "summary", "t, df",
                  "one- or two-tail", "number of pairs", "tukey")


def shared_strings(z):
    """xl/sharedStrings.xml -> [str]  （xlsx 的字符串单元格是索引，必须解析）"""
    try:
        xml = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    except KeyError:
        return []
    return ["".join(RE_SST_T.findall(si.group(1))) for si in RE_SST.finditer(xml)]


def sheet_xml_map(z):
    """sheet 名 -> xl/worksheets/sheetN.xml 的 zip 内路径"""
    wbx = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2tgt = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
        t = m.group(2)
        if not t.startswith("/"):
            t = "xl/" + t.lstrip("./")
        rid2tgt[m.group(1)] = t.replace("xl/xl/", "xl/")
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        tgt = rid2tgt.get(m.group(2))
        if tgt:
            out[m.group(1)] = tgt
    return out


def sig_digits(v):
    """有效小数位数：float 最短往返 repr，去掉尾部 0"""
    s = repr(float(v))
    if "e" in s or "E" in s:
        s = f"{float(v):.17f}".rstrip("0")
    return len(s.split(".")[1].rstrip("0")) if "." in s else 0


def read_sheet(z, path, sst):
    """-> {coord: (raw_str, value, is_formula)}；正确解析 sharedStrings 索引"""
    xml = z.read(path).decode("utf-8", "replace")
    out = {}
    for m in RE_CELL.finditer(xml):
        col, row, attrs, body = m.group(1), int(m.group(2)), m.group(3), m.group(4) or ""
        coord = f"{col}{row}"
        isf = bool(RE_F.search(body))
        ctype = (re.search(r't="([^"]+)"', attrs) or [None, "n"])[1]
        vm = RE_V.search(body)
        tm = RE_T.search(body)
        if ctype == "s" and vm:                     # 共享字符串索引
            try:
                raw = val = sst[int(vm.group(1))]
            except Exception:
                raw = val = vm.group(1)
        elif ctype == "inlineStr" and tm:
            raw = val = tm.group(1)
        elif vm:
            raw = vm.group(1)
            try:
                val = float(raw)
            except ValueError:
                val = raw
        elif tm:
            raw = val = tm.group(1)
        else:
            raw = None; val = None
        if raw is not None:
            out[coord] = (raw, val, isf)
    return out


def load_book(path):
    """-> {sheet: {coord: (raw, value, is_formula)}}"""
    z = zipfile.ZipFile(path)
    smap = sheet_xml_map(z)
    sst = shared_strings(z)
    books = {}
    for name, tgt in smap.items():
        if tgt in z.namelist():
            books[name] = read_sheet(z, tgt, sst)
    z.close()
    return books


# ----------------------------------------------------------------------
# 1. 序列构建：同 sheet、同列/同行、行号连续的数值块
# ----------------------------------------------------------------------
def ws_max_col(cells):
    mx = 1
    for coord in cells:
        m = re.match(r"([A-Z]+)(\d+)$", coord)
        if m:
            mx = max(mx, column_index_from_string(m.group(1)))
    return mx


def label_above(cells, col, first_row):
    """块正上方最近的字符串单元格 —— 用于判定该列是 STAT 统计块还是 DATA 数据列"""
    for r in range(first_row - 1, 0, -1):
        v = cells.get(f"{get_column_letter(col)}{r}")
        if v and isinstance(v[1], str) and v[1].strip():
            return v[1].strip()
    return ""


def classify(items, label):
    """DATA / SEQ / STAT —— SEQ=人工序号列，STAT=GraphPad 派生统计量，二者必须排除"""
    vals = [x[3] for x in items]
    low = (label or "").lower()
    if any(k in low for k in STAT_HEADER_KW):
        return "STAT"
    if len(vals) >= 2 and all(abs(v - i) < 1e-9 for i, v in enumerate(vals, 1)):
        return "SEQ"
    if len(vals) >= 2 and all(abs(v - round(v)) < 1e-9 for v in vals) and \
            len(set(vals)) == len(vals) and max(vals) <= 3:
        return "SEQ"
    return "DATA"


def build_blocks(cells):
    """-> list of dict(orient, key, items=[(coord,row,col,val,raw)], kind, label)
    按【列】分组，再按行号切成连续段 —— 一个段 = 一个数据块。"""
    bycol = defaultdict(list)
    for coord, (raw, val, isf) in cells.items():
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            m = re.match(r"([A-Z]+)(\d+)", coord)
            col, row = column_index_from_string(m.group(1)), int(m.group(2))
            bycol[col].append((coord, row, col, float(val), raw))
    blocks = []
    for col in sorted(bycol):
        items = sorted(bycol[col], key=lambda x: x[1])
        run = [items[0]]
        for it in items[1:]:
            if it[1] - run[-1][1] <= 1:
                run.append(it)
            else:
                blocks.append(_mk_block(col, run, cells))
                run = [it]
        blocks.append(_mk_block(col, run, cells))
    return blocks


def _mk_block(col, run, cells):
    lab = label_above(cells, col, run[0][1])
    b = dict(orient="col", key=col, items=run, label=lab)
    b["kind"] = classify(run, lab)
    return b


def blocks_by_col(blocks):
    d = defaultdict(list)
    for b in blocks:
        if b["orient"] == "col":
            d[b["key"]].append(b)
    return d


# ----------------------------------------------------------------------
# 2. T1 数值指纹探针
# ----------------------------------------------------------------------
def probe_numeric(books, out):
    W = out.write
    W("=" * 100 + "\nT1 数值指纹探针 P1-P8\n" + "=" * 100 + "\n")
    stats = defaultdict(int)

    # P1 逐值全同
    W("\n--- P1 跨序列完全相同的连续值块 (len>=2) ---\n")
    seqs = []
    for fn, sheets in books.items():
        for sname, cells in sheets.items():
            for b in build_blocks(cells):
                if b["kind"] != "DATA":     # 排除人工序号列(SEQ)与 GraphPad 派生统计量(STAT)
                    continue
                seqs.append((f"{fn}::{sname}", b))
    for i in range(len(seqs)):
        for j in range(i + 1, len(seqs)):
            (s1, a), (s2, b) = seqs[i], seqs[j]
            if s1 == s2 and a["orient"] == b["orient"] and a["key"] == b["key"]:
                continue
            va = [x[3] for x in a["items"]]; vb = [x[3] for x in b["items"]]
            n, m = len(va), len(vb)
            L = 0
            while L < min(n, m) and va[L] == vb[L]:
                L += 1
            if L >= 2:
                scope = "CROSS-FILE" if s1 != s2 else "SAME-FILE"
                W(f"[P1 {scope}] {s1} {get_column_letter(a['key'])}{a['items'][0][0][1:]} "
                  f"== {s2} {get_column_letter(b['key'])}{b['items'][0][0][1:]}  len={L}\n")
                W(f"      coords A={[x[0] for x in a['items'][:L]]}\n")
                W(f"      coords B={[x[0] for x in b['items'][:L]]}\n")
                W(f"      values={va[:L]}\n")
                stats["P1"] += 1

    # P2 恒定比值 / P3 恒定偏移（按位置对齐）
    W("\n--- P2 恒定比值 / P3 恒定偏移 (>=3 对, 相对偏差<1e-5) ---\n")
    for i in range(len(seqs)):
        for j in range(i + 1, len(seqs)):
            (s1, a), (s2, b) = seqs[i], seqs[j]
            if a["orient"] != b["orient"]:
                continue
            if s1 == s2 and a["key"] == b["key"]:
                continue
            n = min(len(a["items"]), len(b["items"]))
            if n < 3:
                continue
            va = [x[3] for x in a["items"][:n]]
            vb = [x[3] for x in b["items"][:n]]
            if all(v != 0 for v in vb):
                rat = [va[k] / vb[k] for k in range(n)]
                for tol in (1e-5, 1e-6):
                    groups = defaultdict(list)
                    for k, r in enumerate(rat):
                        groups[round(r, 12)].append(k)
                    for base, idxs in groups.items():
                        if len(idxs) >= 3 and all(abs(rat[k] / base - 1) < tol for k in idxs):
                            frac = Fraction(base).limit_denominator(1000)
                            W(f"[P2 ratio ~{base:.12g} ({frac}) hits={len(idxs)}/{n} tol={tol:g}] "
                              f"{s1} {get_column_letter(a['key'])} vs {s2} {get_column_letter(b['key'])}\n")
                            W(f"      pairs={[(a['items'][k][0], round(va[k],6), b['items'][k][0], round(vb[k],6)) for k in idxs]}\n")
                            stats["P2"] += 1
                            break
                    else:
                        continue
                    break
            diffs = [va[k] - vb[k] for k in range(n)]
            groups = defaultdict(list)
            for k, d in enumerate(diffs):
                groups[round(d, 9)].append(k)
            for d, idxs in groups.items():
                if len(idxs) >= 3 and d != 0:
                    isint = abs(d - round(d)) < 1e-9
                    W(f"[P3 offset {d!r} {'(整数)' if isint else ''} hits={len(idxs)}/{n}] "
                      f"{s1} {get_column_letter(a['key'])} vs {s2} {get_column_letter(b['key'])}\n")
                    W(f"      pairs={[(a['items'][k][0], va[k], b['items'][k][0], vb[k]) for k in idxs]}\n")
                    stats["P3"] += 1

    # P4 等差/等比
    W("\n--- P4 序列内等差/等比 (连续>=4 项) ---\n")
    for sname, b in seqs:
        vals = [x[3] for x in b["items"]]
        coords = [x[0] for x in b["items"]]
        m = len(vals)
        for i in range(m - 3):
            seg = vals[i:i + 4]
            d = seg[1] - seg[0]
            if all(abs(seg[k + 1] - seg[k] - d) < 1e-9 * max(1, abs(d)) for k in range(3)):
                W(f"[P4 arithmetic d={d!r}] {sname} {get_column_letter(b['key'])} "
                  f"coords {coords[i:i+4]} vals {seg}\n")
                stats["P4"] += 1
            if all(v != 0 for v in seg):
                r = seg[1] / seg[0]
                if abs(r - 1) > 1e-12 and all(abs(seg[k + 1] / seg[k] / r - 1) < 1e-9 for k in range(3)):
                    W(f"[P4 geometric r={r:.12g}] {sname} {get_column_letter(b['key'])} "
                      f"coords {coords[i:i+4]} vals {seg}\n")
                    stats["P4"] += 1

    # P5 长尾串（>=8 位尾串 + >=13 位有效数字）
    W("\n--- P5 长十进制尾串共享 (尾串>=8位 且 有效位>=13) ---\n")
    index = defaultdict(list)
    for sname, b in seqs:
        for coord, row, col, val, raw in b["items"]:
            if sig_digits(val) >= 13:
                s = repr(val)
                frac = s.split(".")[1] if "." in s else ""
                if len(frac) >= 8:
                    index[frac[-8:]].append((sname, coord, val, raw))
    for tail, lst in sorted(index.items(), key=lambda kv: -len(kv[1])):
        panels = {s for s, _, _, _ in lst}
        if len(panels) >= 2 or len(lst) >= 3:
            W(f"[P5 tail .{tail}] x{len(lst)} panels={sorted(panels)}\n")
            for s, c, v, raw in lst:
                W(f"      {s} {c} = {raw}\n")
            stats["P5"] += 1

    # P6 组内重复值
    W("\n--- P6 同一组内完全重复的值 (n<=12) ---\n")
    for sname, b in seqs:
        if not (2 <= len(b["items"]) <= 12):
            continue
        seen = defaultdict(list)
        for coord, row, col, val, raw in b["items"]:
            seen[repr(val)].append((coord, raw))
        for v, lst in seen.items():
            if len(lst) >= 2:
                W(f"[P6 dup {v} x{len(lst)}] {sname} {get_column_letter(b['key'])} {lst}\n")
                stats["P6"] += 1

    # P7 精度断层
    W("\n--- P7 同 sheet 内精度断层 (有效位差>=3) ---\n")
    for fn, sheets in books.items():
      for sname, cells in sheets.items():
        per = {}
        for b in build_blocks(cells):
            if len(b["items"]) < 2 or b["kind"] != "DATA":
                continue
            sd = [sig_digits(x[3]) for x in b["items"]]
            per[b["key"]] = (max(sd), min(sd),
                             [(x[0], x[3], sd[i]) for i, x in enumerate(b["items"])])
        vals = list(per.values())
        for i in range(len(vals)):
            for j in range(i + 1, len(vals)):
                a, b = vals[i], vals[j]
                if abs(a[0] - b[1]) >= 3 or abs(b[0] - a[1]) >= 3:
                    W(f"[P7 precision break] {sname} col {get_column_letter(list(per)[i])} "
                      f"(max {a[0]}/min {a[1]}) vs col {get_column_letter(list(per)[j])} "
                      f"(max {b[0]}/min {b[1]})\n")
                    W(f"      A={a[2][:6]}\n      B={b[2][:6]}\n")
                    stats["P7"] += 1

    # P8 粒度网格
    W("\n--- P8 值落在 1/5 或 1/19 等粒度网格 ---\n")
    for sname, b in seqs:
        for coord, row, col, val, raw in b["items"]:
            for den in (5, 19):
                if abs(val * den - round(val * den)) < 1e-6 and abs(val * den) >= 10:
                    W(f"[P8 grid 1/{den}] {sname} {coord} = {raw}\n")
                    stats["P8"] += 1
                    break

    W(f"\n>>> 命中计数: {dict(stats)}\n")
    return stats


# ----------------------------------------------------------------------
# 3. T2 低样本 n<=3 专项
# ----------------------------------------------------------------------
def probe_lown(books, out):
    W = out.write
    W("=" * 100 + "\nT2 低样本 (n<=3) 专项 L1-L5\n" + "=" * 100 + "\n")
    stats = defaultdict(int)
    for fn, sheets in books.items():
      for sname, cells in sheets.items():
        blocks = build_blocks(cells)
        # 低样本组：2~3 个值；单值块与序号/统计块一律排除（否则产生海量假阳性）
        small = [b for b in blocks if 2 <= len(b["items"]) <= 3 and b["kind"] == "DATA"]
        if not small:
            continue
        W(f"\n### {sname}  (n<=3 组 {len(small)} 个)\n")
        for i in range(len(small)):
            for j in range(i + 1, len(small)):
                a, b = small[i], small[j]
                va = [x[3] for x in a["items"]]; vb = [x[3] for x in b["items"]]
                ca = [x[0] for x in a["items"]]; cb = [x[0] for x in b["items"]]
                # L1 首值相同（跳过 1.0 归一化锚点）
                if va[0] == vb[0] and not (va[0] == 1.0 and vb[0] == 1.0):
                    W(f"[L1 首值相同] {sname}: {ca[0]}={a['items'][0][4]} vs {cb[0]}={b['items'][0][4]}\n")
                    stats["L1"] += 1
                # L2 精确整数偏移
                offs = [va[k] - vb[k] for k in range(min(len(va), len(vb)))]
                if len(offs) >= 2:
                    same = [o for o in offs if abs(o - offs[0]) < 1e-12]
                    if len(same) >= 2 and abs(offs[0] - round(offs[0])) < 1e-12 and offs[0] != 0:
                        W(f"[L2 恒定整数偏移 {offs[0]!r}] {sname}: {ca} vs {cb}  offsets={offs}\n")
                        stats["L2"] += 1
                # L5 逐值全同
                if len(va) == len(vb) and va == vb:
                    W(f"[L5 逐值全同] {sname}: {ca} vs {cb}  values={va}\n")
                    stats["L5"] += 1
        # L3 精度断层：块级有效位数（取块内最小值），差 >=3 才报；两块都需 >=2 值
        sds = []
        for b in small:
            sdv = [sig_digits(x[3]) for x in b["items"]]
            sds.append((b["key"], min(sdv), max(sdv),
                        [(x[0], x[3]) for x in b["items"]]))
        for i in range(len(sds)):
            for j in range(i + 1, len(sds)):
                if sds[i][0] == sds[j][0]:
                    continue
                gap = max(abs(sds[i][1] - sds[j][1]), abs(sds[i][2] - sds[j][2]))
                if gap >= 3:
                    W(f"[L3 精度断层 gap={gap}] {sname}: col {get_column_letter(sds[i][0])} "
                      f"(sig {sds[i][1]}-{sds[i][2]}) vs col {get_column_letter(sds[j][0])} "
                      f"(sig {sds[j][1]}-{sds[j][2]})\n")
                    W(f"      A={sds[i][3]}\n      B={sds[j][3]}\n")
                    stats["L3"] += 1
    W(f"\n>>> 命中计数: {dict(stats)}\n")
    return stats


# ----------------------------------------------------------------------
# 4. T0 元数据 / 容器取证
# ----------------------------------------------------------------------
def probe_meta(paths, out):
    W = out.write
    W("=" * 100 + "\nT0 元数据与容器取证\n" + "=" * 100 + "\n")
    rows = []
    for p in paths:
        fn = os.path.basename(p)
        z = zipfile.ZipFile(p)
        def rd(name):
            try:
                return z.read(name).decode("utf-8", "replace")
            except KeyError:
                return ""
        core = rd("docProps/core.xml")
        app = rd("docProps/app.xml")
        wbx = rd("xl/workbook.xml")
        creator = (re.search(r"<dc:creator>(.*?)</dc:creator>", core) or [None, ""])[1]
        lastmod = (re.search(r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>", core) or [None, ""])[1]
        created = (re.search(r"<dcterms:created[^>]*>(.*?)</dcterms:created>", core) or [None, ""])[1]
        modified = (re.search(r"<dcterms:modified[^>]*>(.*?)</dcterms:modified>", core) or [None, ""])[1]
        company = (re.search(r"<Company>(.*?)</Company>", app) or [None, ""])[1]
        appver = (re.search(r"<AppVersion>(.*?)</AppVersion>", app) or [None, ""])[1]
        abs_path = (re.search(r"absPath[^>]*url=\"([^\"]*)\"", wbx) or [None, ""])[1]
        ztimes = sorted({i.date_time for i in z.infolist()})
        z.close()
        rows.append(dict(file=fn, creator=creator, lastModifiedBy=lastmod,
                         created=created, modified=modified, company=company,
                         appversion=appver, absPath=abs_path,
                         zip_first=("%04d-%02d-%02d %02d:%02d:%02d" % ztimes[0]) if ztimes else "",
                         zip_distinct=len(ztimes)))
        W(f"\n### {fn}\n")
        W(f"  creator        = {creator!r}\n  lastModifiedBy = {lastmod!r}\n")
        W(f"  created        = {created}\n  modified       = {modified}\n")
        W(f"  Company        = {company!r}   AppVersion = {appver!r}\n")
        W(f"  x15ac:absPath  = {abs_path!r}\n")
        W(f"  zip 内嵌时间戳  : 首={rows[-1]['zip_first']}  不同值个数={rows[-1]['zip_distinct']}"
          f"{'   <== 全部 1980-01-01：容器被非 Excel 工具重打包' if rows[-1]['zip_distinct']<=2 and str(ztimes[0][0])=='1980' else ''}\n")
    W("\n--- 汇总表 ---\n")
    W(f"{'file':<16}{'creator':<12}{'lastModifiedBy':<16}{'created':<26}{'modified':<26}{'absPath'}\n")
    for r in rows:
        W(f"{r['file']:<16}{r['creator']:<12}{r['lastModifiedBy']:<16}{r['created']:<26}{r['modified']:<26}{r['absPath']}\n")
    W("\n>>> 同秒 created 聚簇检测（批量生成的危险信号）：\n")
    cc = defaultdict(list)
    for r in rows:
        cc[r["created"]].append(r["file"])
    for c, fs in cc.items():
        tag = "  <== 全部同秒：疑似批量生成（需结合 modified 是否分散判断）" if len(fs) >= 5 else ""
        W(f"    created={c} -> {len(fs)} 个文件{tag}\n")
    W(">>> modified 同秒聚簇检测（真批量生成才会出现）：\n")
    mc = defaultdict(list)
    for r in rows:
        mc[r["modified"]].append(r["file"])
    for c, fs in mc.items():
        if len(fs) >= 3:
            W(f"    modified={c} -> {len(fs)} 个文件  <== 同秒修改\n")
    return rows


# ----------------------------------------------------------------------
# 5. T2 统计量自洽性复算
# ----------------------------------------------------------------------
def probe_stats(books, out, control_kw=None):
    from scipy import stats as sps
    W = out.write
    W("=" * 100 + "\nT2 统计量自洽性复算（Mean Diff. / t）\n" + "=" * 100 + "\n")
    W("纪律: ① 必须设已知干净面板作对照；② t 检验优先于均值差；③ 按存储显示精度归一比较\n")
    rows = []
    for fn, sheets in books.items():
      for sname, cells in sheets.items():
        # 找 GraphPad 统计块：含 "mean diff." 的表头行
        hdr = {}
        for coord, (raw, val, isf) in cells.items():
            if isinstance(val, str):
                hdr[coord] = val
        mean_diff_hdr = [(c, v) for c, v in hdr.items() if "mean diff" in v.lower()]
        t_hdr = [(c, v) for c, v in hdr.items() if v.strip().lower() == "t, df"]
        W(f"[scan] {fn}::{sname}  字符串格={len(hdr)}  meanDiff表头={mean_diff_hdr}  t表头={t_hdr}\n")
        if not mean_diff_hdr and not t_hdr:
            continue
        # 数据列：同行其它表头 + 有 >=2 数值
        def header_cols(row):
            cols = {}
            for c, v in hdr.items():
                m = re.match(r"([A-Z]+)(\d+)$", c)
                if m and int(m.group(2)) == row and v.strip():
                    cols[column_index_from_string(m.group(1))] = v.strip()
            return cols
        # --- t 检验块 ---
        for thc, _ in t_hdr:
            tm = re.match(r"([A-Z]+)(\d+)$", thc)
            trow, tcol = int(tm.group(2)), column_index_from_string(tm.group(1))
            # 数据列：从块往上走【连续数值行】，其上一行即该组的数据表头行
            # （GraphPad 的 t 块里 'One- or two-tailed'/'Two-tailed' 也是表头，不能选它）
            runs = {}
            for c in range(1, ws_max_col(cells) + 1):
                allrows = []
                for r in range(1, trow):
                    cc = f"{get_column_letter(c)}{r}"
                    if cc in cells:
                        raw, val, isf = cells[cc]
                        if isinstance(val, (int, float)) and not isinstance(val, bool):
                            allrows.append(r)
                if not allrows:
                    continue
                # 取末尾的连续组（允许中间夹 1-2 行空白）
                allrows.sort()
                grp = [allrows[-1]]
                for r in reversed(allrows[:-1]):
                    if grp[-1] - r <= 3:
                        grp.append(r)
                    else:
                        break
                if len(grp) >= 2:
                    grp = sorted(grp)
                    # 只排除【无表头】的 1..n 序号列；有表头的整数数据列是真实测量，必须保留
                    h = hdr.get(f"{get_column_letter(c)}{min(grp)-1}", "").strip()
                    vals = [float(cells[f"{get_column_letter(c)}{rr}"][1]) for rr in grp]
                    is_seq = (not h) and all(abs(v - i) < 1e-9 for i, v in enumerate(vals, 1))
                    if not is_seq:
                        runs[c] = (sorted(grp), min(grp) - 1)
            groups = defaultdict(list)
            for c, (rs, hr) in runs.items():
                groups[hr].append(c)
            if not groups:
                continue
            best_hr = max(groups, key=lambda k: len(groups[k]))
            data_cols = {}
            for c in groups[best_hr]:
                h = hdr.get(f"{get_column_letter(c)}{best_hr}", f"col{c}")
                data_cols[c] = h.strip()
            cand = [(c, data_cols[c], [float(cells[f"{get_column_letter(c)}{rr}"][1])
                                      for rr in runs[c][0]]) for c in groups[best_hr]]
            # 找该块的 P value 单元格与 t 值单元格（整行/邻近行搜索，不限列）
            #   GraphPad 版式：E7='t, df'（表头）  F7='t=11.26, df=2'（值）—— 二者不同列
            pcell = None
            for r in range(max(1, trow - 10), trow + 12):
                for cc, vv in hdr.items():
                    if re.match(rf"[A-Z]+{r}$", cc) and vv.strip().lower() == "p value":
                        pcell = cc; break
                if pcell: break
            tv = None
            for r in range(trow, trow + 10):
                for cc, vv in hdr.items():
                    if re.match(rf"[A-Z]+{r}$", cc) and re.search(r"t=-?[\d.]+\s*,\s*df=", vv):
                        tv = vv; break
                if tv: break
            if tv is None:
                continue
            tt = float(re.search(r"t=(-?[\d.]+)", tv).group(1))
            df_stored = int(re.search(r"df=(\d+)", tv).group(1))
            if len(cand) < 2:
                continue
            # 取紧邻的前两列（按列序）；候选多于 2 时不硬判，列出候选交人工核对
            if len(cand) > 2:
                W(f"\n[t-test] {sname}  候选列多于 2，**需人工指定配对**\n")
                W(f"   stored: {tv}   P={hdr.get(pcell,'?')}\n")
                W(f"   候选: {[(c, h, len(v), v) for c, h, v in cand]}\n")
                rows.append(dict(kind="t", sheet=sname, rel=0.0, stored=tv,
                                 recomp="需人工配对"))
                continue
            (c1, h1, v1), (c2, h2, v2) = cand[0], cand[1]
            # 区分配对 / 非配对 t 检验（GraphPad 的块标题里写着 "Paired t test" / "Unpaired t test"）
            all_hdr_txt = " ".join(v.lower() for v in hdr.values())
            paired = ("paired t test" in all_hdr_txt) and ("unpaired t test" not in all_hdr_txt)
            try:
                if paired and len(v1) == len(v2):
                    t, p = sps.ttest_rel(v1, v2)
                else:
                    t, p = sps.ttest_ind(v1, v2, equal_var=True)
            except Exception:
                continue
            rel = abs(t - tt) / max(abs(tt), 1e-9)
            W(f"\n[t-test] {sname}  {h1} vs {h2}\n")
            W(f"   stored: {tv}   P={hdr.get(pcell,'?')}\n")
            W(f"   recomp: t={t:.6f}  P={p:.6f}  ({'配对' if paired else '非配对'} t, n1={len(v1)}, n2={len(v2)}, df={df_stored})\n")
            W(f"   -> |t_recomp - t_stored| = {abs(t-tt):.6f}  (相对 {rel*100:.3f}%)  "
              f"{'不一致' if rel > 0.02 else '一致'}\n")
            rows.append(dict(kind="t", sheet=sname, rel=rel, stored=tv,
                             recomp=f"t={t:.6f} P={p:.6f}"))
    # --- Mean Diff. 块 ---
        for mhc, _ in mean_diff_hdr:
            mh = re.match(r"([A-Z]+)(\d+)$", mhc)
            hrow = int(mh.group(2))
            label_col = hrow
            # 标签列 = 同一行 "Tukey" 或 "Unpaired t"/"Paired t" 所在列
            lab_col = None
            for c, v in sorted(hdr.items()):
                m = re.match(r"([A-Z]+)(\d+)$", c)
                if m and int(m.group(2)) == hrow and ("tukey" in v.lower() or "comparisons" in v.lower()):
                    lab_col = column_index_from_string(m.group(1)); break
            if lab_col is None:
                continue
            md_col = column_index_from_string(mh.group(1))
            # 数据表头行：同一行内 label 列左边的表头
            data_cols = {c: v for c, v in header_cols(hrow).items() if c < lab_col}
            # 若本行没有表头（统计块在下方），向上找最近的表头行
            if len(data_cols) < 2:
                for hr in range(hrow - 1, max(0, hrow - 4), -1):
                    dc = {c: v for c, v in header_cols(hr).items() if c < lab_col}
                    if len(dc) >= 2:
                        data_cols = dc; break
            if len(data_cols) < 2:
                continue
            # 数据行区间
            def num_rows(col):
                """统计块**下方**该列的数值行（数据行，不是表头行）"""
                rs = []
                for r in range(hrow + 1, hrow + 80):
                    cc = f"{get_column_letter(col)}{r}"
                    if cc in cells:
                        raw, val, isf = cells[cc]
                        if isinstance(val, (int, float)) and not isinstance(val, bool):
                            rs.append(r)
                return rs
            cols_with_data = {}
            for c in data_cols:
                rs = num_rows(c)
                if len(rs) >= 2:
                    cols_with_data[c] = rs
            if len(cols_with_data) < 2:
                continue
            for r in range(hrow + 1, hrow + 60):
                lbl = hdr.get(f"{get_column_letter(lab_col)}{r}")
                sv = cells.get(f"{get_column_letter(md_col)}{r}")
                if lbl is None or sv is None:
                    continue
                if not isinstance(sv[1], (int, float)):
                    continue
                m = re.match(r"^(.*?)\s+vs\.?\s+(.*)$", lbl.strip())
                if not m:
                    continue
                def find_col(name):
                    nm = re.sub(r"\s+", " ", name).strip().lower()
                    for c, h in data_cols.items():
                        if re.sub(r"\s+", " ", h).strip().lower() == nm and c in cols_with_data:
                            return c
                    return None
                c1, c2 = find_col(m.group(1)), find_col(m.group(2))
                if c1 is None or c2 is None:
                    continue
                v1 = [float(cells[f"{get_column_letter(c1)}{rr}"][1]) for rr in cols_with_data[c1]]
                v2 = [float(cells[f"{get_column_letter(c2)}{rr}"][1]) for rr in cols_with_data[c2]]
                d = float(np.mean(v1) - np.mean(v2))
                stored = float(sv[1])
                sstr = sv[0]
                dec = len(sstr.split(".")[1].rstrip("0")) if "." in sstr else 0
                tol = (0.5 * 10 ** (-dec) + 1e-12) if dec else 5e-4
                delta = round(d, dec) - stored if dec else d - stored
                ok = abs(delta) <= tol
                W(f"[MeanDiff] {sname} r{r} {lbl.strip()[:44]:46s} stored={stored:<14} recomp={d:.8f} "
                  f"delta={delta:+.6f} {'OK' if ok else 'DIFF'}\n")
                rows.append(dict(kind="md", sheet=sname, rel=abs(delta) / max(abs(stored), 1.0),
                                 stored=str(stored), recomp=f"{d:.8f}", ok=ok))
    W("\n>>> 汇总\n")
    if rows:
        rels = [r["rel"] for r in rows]
        W(f"    条目 {len(rows)}  相对偏差中位 {np.median(rels):.6f}  最大 {max(rels):.6f}\n")
        if control_kw:
            ctl = [r for r in rows if control_kw.lower() in r["sheet"].lower()]
            oth = [r for r in rows if control_kw.lower() not in r["sheet"].lower()]
            if ctl:
                W(f"    对照面板 '{control_kw}': {len(ctl)} 条, 相对偏差中位 {np.median([r['rel'] for r in ctl]):.6f}, "
                  f"最大 {max(r['rel'] for r in ctl):.6f}\n")
            if oth:
                W(f"    其余面板      : {len(oth)} 条, 相对偏差中位 {np.median([r['rel'] for r in oth]):.6f}, "
                  f"最大 {max(r['rel'] for r in oth):.6f}\n")
            if ctl and oth:
                W("    => 只有当其余面板的偏差**显著超过**对照面板时，才可判为『数据与统计不自洽』。\n")
    return rows


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--control", default=None)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    dumpdir = os.path.join(a.out, "dump")
    os.makedirs(dumpdir, exist_ok=True)

    paths = sorted([os.path.join(a.input, f) for f in os.listdir(a.input)
                    if f.lower().endswith(".xlsx")])
    if not paths:
        print("no xlsx found in", a.input); return 1
    print(f"[sci-audit] {len(paths)} 个 xlsx")

    books = {}
    with open(os.path.join(dumpdir, "cells.jsonl"), "w", encoding="utf-8") as jf:
        for p in paths:
            fn = os.path.basename(p)
            try:
                bk = load_book(p)
            except Exception as e:
                print(f"  !! {fn}: {e}"); continue
            books[fn] = bk
            with open(os.path.join(dumpdir, fn.replace(".xlsx", "") + "_dump.txt"),
                      "w", encoding="utf-8") as tf:
                for sname, cells in bk.items():
                    tf.write(f"\n===== {fn} / {sname} =====\n")
                    for coord in sorted(cells, key=lambda c: (int(re.match(r'[A-Z]+(\d+)', c).group(1)),
                                                               column_index_from_string(re.match(r'([A-Z]+)', c).group(1)))):
                        raw, val, isf = cells[coord]
                        tf.write(f"{coord}\t{raw}\t{'FORMULA' if isf else ''}\n")
            for sname, cells in bk.items():
                for coord, (raw, val, isf) in cells.items():
                    jf.write(json.dumps({"file": fn, "sheet": sname, "coord": coord,
                                         "raw": raw, "value": val, "formula": isf},
                                        ensure_ascii=False) + "\n")
            print(f"  ok {fn}: {len(bk)} sheets")

    with open(os.path.join(a.out, "findings_numeric.txt"), "w", encoding="utf-8") as f:
        probe_numeric(books, f)
    with open(os.path.join(a.out, "findings_lown.txt"), "w", encoding="utf-8") as f:
        probe_lown(books, f)
    with open(os.path.join(a.out, "findings_meta.txt"), "w", encoding="utf-8") as f:
        probe_meta(paths, f)
    with open(os.path.join(a.out, "findings_stats.txt"), "w", encoding="utf-8") as f:
        probe_stats(books, f, a.control)

    # summary.md
    def head(p, n=40):
        try:
            with open(os.path.join(a.out, p), encoding="utf-8") as f:
                return "".join([next(f) for _ in range(n)])
        except StopIteration:
            return open(os.path.join(a.out, p), encoding="utf-8").read()
        except Exception:
            return ""
    sm = ["# sci-audit 探针摘要\n"]
    for name, p in [("数值指纹 P1-P8", "findings_numeric.txt"),
                    ("低样本 n<=3 L1-L5", "findings_lown.txt"),
                    ("元数据/容器", "findings_meta.txt"),
                    ("统计自洽性", "findings_stats.txt")]:
        sm.append(f"\n## {name}  ({p})\n\n```\n" + head(p, 60) + "\n```\n")
    with open(os.path.join(a.out, "summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(sm))
    print(f"[sci-audit] 输出目录: {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
