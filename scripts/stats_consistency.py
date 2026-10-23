# -*- coding: utf-8 -*-
"""
统计量自洽性复核（主 agent 版，v2）
------------------------------------------------------------------
判据要点：GraphPad 统计块存的是**四舍五入后的显示值**
          (Mean Diff. '9.313'、'0.2258'；CI '6.394 to 12.23')
          拿全精度重算值直接比对 -> 必然出现 1e-4~1e-3 伪 delta。
          正确做法：把重算值按**与存储串相同的小数位数**取整后再比。

布局两种，均支持：
  A. 统计块在数据块下方（label 列=1, mean diff 列=2）
  B. 统计块在数据块右侧（label 列=cTukey, mean diff 列=cMean）
输出：_dump/stats_consistency.txt (UTF-8)
"""
import openpyxl, os, re
from openpyxl.utils import get_column_letter
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "_dump", "stats_consistency.txt")
lines = []
def w(s=""): lines.append(s)

def dec_of(s):
    s = str(s).strip()
    if "." in s:
        return len(s.split(".")[1].rstrip("0"))
    return 0

def norm(s): return re.sub(r"\s+", " ", str(s)).strip().lower()

def is_num(v): return isinstance(v, (int, float)) and not isinstance(v, bool)

w("=" * 104)
w("统计量自洽性复核 v2 —— 按【与存储串相同精度】比对")
w("=" * 104)

grand_ok = grand_n = 0
per_panel = []

for n in range(7, 17):
    path = os.path.join(BASE, f"MOESM{n}.xlsx")
    if not os.path.exists(path): continue
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        # ---- 定位统计块头 ----
        blk = None
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and "tukey" in v.lower() and "multiple comparison" in v.lower():
                    # 同行右侧找 "Mean Diff"
                    cm = None
                    for c2 in range(c + 1, ws.max_column + 1):
                        v2 = ws.cell(r, c2).value
                        if isinstance(v2, str) and "mean diff" in v2.lower():
                            cm = c2; break
                    if cm: blk = (r, c, cm); break
            if blk: break
        if not blk:
            wb.close(); continue
        br, clab, cmean = blk

        # ---- 定位数据表头行 ----
        def strheaders(row, cmax):
            cols = []
            for c in range(1, cmax + 1):
                h = ws.cell(row, c).value
                if isinstance(h, str) and h.strip():
                    cols.append((c, h.strip()))
            return cols
        hdr_cands = []
        for row in range(1, br + 1):
            cols = strheaders(row, clab - 1 if clab > 1 else ws.max_column)
            # 只认"看起来像数据表头"：>=2 列且多为非统计词
            if len(cols) >= 2:
                hdr_cands.append((row, cols))
        if not hdr_cands:
            wb.close(); continue
        hdr_row, hcols = hdr_cands[-1] if clab > 1 else hdr_cands[0]
        # B 布局：表头行 == br；A 布局：取 br 之前最近的有 >=2 表头的行
        if clab > 1:
            cand = [(rw, cs) for rw, cs in hdr_cands if rw == br and len(cs) >= 2]
            if cand: hdr_row, hcols = cand[0]
            else: hdr_row, hcols = hdr_cands[-1]
        else:
            hdr_row, hcols = hdr_cands[0]

        # ---- 数据行：hdr_row 之后连续含数值的行 ----
        d0 = hdr_row + 1
        d1 = d0 - 1
        r = d0
        while r <= ws.max_row:
            if any(is_num(ws.cell(r, c).value) for c, _ in hcols):
                d1 = r; r += 1
            else:
                break
        if d1 < d0:
            wb.close(); continue

        # 数据列 = 表头列中实际有 >=2 个数值的
        dcols = []
        for c, h in hcols:
            vals = [ws.cell(x, c).value for x in range(d0, d1 + 1)]
            if sum(1 for v in vals if is_num(v)) >= 2:
                dcols.append((c, h))

        w("-" * 104)
        w(f"### MOESM{n}.xlsx / '{ws.title}'   块头行={br} 标签列={get_column_letter(clab)} MeanDiff列={get_column_letter(cmean)}")
        w(f"    数据表头行={hdr_row}  数据行={d0}..{d1}  数据列={[(get_column_letter(c), h) for c, h in dcols]}")
        ok = bad = skipped = 0
        for r in range(br + 1, ws.max_row + 1):
            label = ws.cell(r, clab).value
            stored = ws.cell(r, cmean).value
            if not isinstance(label, str) or not label.strip():
                if ok + bad > 0 and r > br + 1: break
                continue
            if stored is None or not is_num(stored): continue
            m = re.match(r"^(.*?)\s+vs\.?\s+(.*)$", label.strip())
            if not m: continue
            g1, g2 = norm(m.group(1)), norm(m.group(2))
            c1 = next((c for c, h in dcols if norm(h) == g1), None)
            c2 = next((c for c, h in dcols if norm(h) == g2), None)
            if c1 is None or c2 is None:
                skipped += 1
                w(f"    r{r}: {label[:46]:48s} [表头未匹配: {g1!r} / {g2!r}]")
                continue
            v1 = [v for v in (ws.cell(x, c1).value for x in range(d0, d1 + 1)) if is_num(v)]
            v2 = [v for v in (ws.cell(x, c2).value for x in range(d0, d1 + 1)) if is_num(v)]
            if not v1 or not v2: continue
            diff = float(np.mean(v1)) - float(np.mean(v2))
            dec = dec_of(stored)
            r_ = round(diff, dec) if dec else float(f"{diff:.4g}")
            hit = abs(r_ - float(stored)) <= (0.5 * 10 ** (-dec) + 1e-12) if dec else abs(r_ - float(stored)) <= 5e-4
            if hit: ok += 1
            else:
                bad += 1
                w(f"    r{r}: {label[:44]:46s} stored={stored!s:>11s} recomputed={diff:.10f} @{dec}dp={r_}  MISMATCH")
        grand_ok += ok; grand_n += ok + bad
        w(f"  ==> Mean Diff. 自洽 {ok}/{ok+bad}" + (f"  (跳过 {skipped} 条标签未匹配)" if skipped else ""))
        per_panel.append((f"MOESM{n} {ws.title}", ok, ok + bad))
    wb.close()

w("")
w("=" * 104)
w(f"TOTAL Mean Diff. 自洽: {grand_ok}/{grand_n}  ({grand_ok/max(grand_n,1)*100:.1f}%)")
w("=" * 104)
open(OUT, "w", encoding="utf-8").write("\n".join(lines))
print(f"Mean Diff self-consistency: {grand_ok}/{grand_n}")
