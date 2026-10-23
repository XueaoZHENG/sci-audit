# Unified blind numeric detectors over _dump/cells.jsonl
# Output: _dump/findings.txt (UTF-8, Chinese OK)
import json, os, math
from collections import defaultdict
from fractions import Fraction

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DUMP = os.path.join(BASE, "_dump")

cells = []
with open(os.path.join(DUMP, "cells.jsonl"), encoding="ascii") as f:
    for line in f:
        rec = json.loads(line)
        cells.append(rec)

def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

# numeric cells indexed per (file, sheet)
by_sheet = defaultdict(list)
for rec in cells:
    v = rec["value"]
    if isinstance(v, str):
        try:
            v = float(v)
        except Exception:
            continue
    if is_num(v):
        by_sheet[(rec["file"], rec["sheet"])].append(
            (rec["row"], rec["col"], rec["coord"], float(v)))

# Build sequences: contiguous column runs and row runs (gap<=1 row/col, length>=2)
def build_seqs():
    seqs = []  # each: dict(file, sheet, orient, key, items=[(coord, row, col, val)])
    for (fn, sn), lst in by_sheet.items():
        cols = defaultdict(list)
        rows = defaultdict(list)
        for r, c, coord, v in lst:
            cols[c].append((r, coord, v))
            rows[r].append((c, coord, v))
        for c, items in cols.items():
            items.sort()
            # split into contiguous runs
            run = []
            prev = None
            for r, coord, v in items:
                if prev is not None and r - prev > 1:
                    if len(run) >= 2:
                        seqs.append({"file": fn, "sheet": sn, "orient": "col", "key": c, "items": run})
                    run = []
                run.append((coord, r, c, v))
                prev = r
            if len(run) >= 2:
                seqs.append({"file": fn, "sheet": sn, "orient": "col", "key": c, "items": run})
        for r, items in rows.items():
            items.sort()
            run = []
            prev = None
            for c, coord, v in items:
                if prev is not None and c - prev > 1:
                    if len(run) >= 2:
                        seqs.append({"file": fn, "sheet": sn, "orient": "row", "key": r, "items": run})
                    run = []
                run.append((coord, r, c, v))
                prev = c
            if len(run) >= 2:
                seqs.append({"file": fn, "sheet": sn, "orient": "row", "key": r, "items": run})
    return seqs

seqs = build_seqs()

def sig_decimals(v):
    s = repr(v)
    if "." in s:
        return len(s.split(".")[1].rstrip("0"))
    return 0

def interesting(vals):
    # at least one value with >=4 decimal places or |v|>=1000 -> not trivial small-int noise
    return any(sig_decimals(v) >= 4 or abs(v) >= 1000 for v in vals)

out = open(os.path.join(DUMP, "findings.txt"), "w", encoding="utf-8")
def w(s=""):
    out.write(s + "\n")

def label(seq):
    o = seq["orient"]
    k = seq["key"]
    if o == "col":
        from openpyxl.utils import get_column_letter
        kl = get_column_letter(k)
    else:
        kl = str(k)
    return f'{seq["file"]} [{seq["sheet"]}] {o} {kl} rows/pos {seq["items"][0][0]}..{seq["items"][-1][0]}'

# ---------- Detector 2 & 3: identical contiguous runs between sequence pairs ----------
w("##### DETECTOR 2/3: identical contiguous value runs (len>=2), exact and round4 #####")
def common_runs(va, vb, keyfn):
    # longest common contiguous subsequences of values, len>=2
    va = [keyfn(x) for x in va]
    vb = [keyfn(x) for x in vb]
    hits = []
    n, m = len(va), len(vb)
    # simple DP for all matches len>=2 (start positions)
    for i in range(n):
        for j in range(m):
            L = 0
            while i + L < n and j + L < m and va[i+L] == vb[j+L]:
                L += 1
            if L >= 2:
                hits.append((i, j, L))
    # keep only maximal (not contained in a longer hit at shifted pos)
    maximal = []
    for (i, j, L) in hits:
        contained = False
        for (i2, j2, L2) in hits:
            if (i, j, L) != (i2, j2, L2) and i2 <= i and j2 <= j and i2+L2 >= i+L and j2+L2 >= j+L and L2 > L:
                contained = True
                break
        if not contained:
            maximal.append((i, j, L))
    return maximal

pairs_reported = 0
for si in range(len(seqs)):
    for sj in range(si+1, len(seqs)):
        A, B = seqs[si], seqs[sj]
        same_file = A["file"] == B["file"]
        same_seq = same_file and A["sheet"] == B["sheet"] and A["orient"] == B["orient"] and A["key"] == B["key"]
        if same_seq:
            continue
        for mode, keyfn in (("exact", lambda v: v), ("round4", lambda v: round(v, 4))):
            hits = common_runs([x[3] for x in A["items"]], [x[3] for x in B["items"]], keyfn)
            for (i, j, L) in hits:
                vals = [x[3] for x in A["items"][i:i+L]]
                if not interesting(vals):
                    continue
                scope = "SAME-FILE" if same_file else "CROSS-FILE"
                w(f"[{scope} identical run len={L} mode={mode}]")
                w(f"  A: {label(A)} coords {[x[0] for x in A['items'][i:i+L]]}")
                w(f"  B: {label(B)} coords {[x[0] for x in B['items'][j:j+L]]}")
                w(f"  values: {vals}")
                pairs_reported += 1
w(f"total runs reported: {pairs_reported}\n")

# ---------- Detector 4: constant offset ----------
w("##### DETECTOR 4: constant offset between aligned sequences (>=3 pairs) #####")
# align sequences by row overlap (col-oriented) or by position index
def aligned_pairs(A, B):
    # align by absolute row if both col-oriented in same sheet, else by index
    if A["orient"] == "col" and B["orient"] == "col" and A["sheet"] == B["sheet"] and A["file"] == B["file"]:
        ra = {x[1]: x for x in A["items"]}
        rb = {x[1]: x for x in B["items"]}
        rows = sorted(set(ra) & set(rb))
        return [(ra[r], rb[r]) for r in rows]
    else:
        n = min(len(A["items"]), len(B["items"]))
        return [(A["items"][k], B["items"][k]) for k in range(n)]

cnt4 = 0
for si in range(len(seqs)):
    for sj in range(si+1, len(seqs)):
        A, B = seqs[si], seqs[sj]
        if A["orient"] != B["orient"]:
            continue
        pairs = aligned_pairs(A, B)
        if len(pairs) < 3:
            continue
        diffs = [round(pa[3] - pb[3], 9) for (pa, pb) in pairs]
        groups = defaultdict(list)
        for k, d in enumerate(diffs):
            groups[d].append(k)
        for d, idxs in groups.items():
            if len(idxs) >= 3 and d != 0:
                vals_a = [pairs[k][0][3] for k in idxs]
                if not interesting(vals_a):
                    continue
                scope = "SAME-FILE" if A["file"] == B["file"] else "CROSS-FILE"
                w(f"[{scope} constant offset {d} hits={len(idxs)}/{len(pairs)}]")
                w(f"  A: {label(A)}")
                w(f"  B: {label(B)}")
                w(f"  pairs: {[(pairs[k][0][0], pairs[k][0][3], pairs[k][1][0], pairs[k][1][3]) for k in idxs]}")
                cnt4 += 1
w(f"total offset hits: {cnt4}\n")

# ---------- Detector 5: constant ratio ----------
w("##### DETECTOR 5: constant ratio between aligned sequences (>=3 pairs, rel dev<1e-6) #####")
cnt5 = 0
for si in range(len(seqs)):
    for sj in range(si+1, len(seqs)):
        A, B = seqs[si], seqs[sj]
        if A["orient"] != B["orient"]:
            continue
        pairs = aligned_pairs(A, B)
        if len(pairs) < 3:
            continue
        ratios = []
        ok = True
        for (pa, pb) in pairs:
            if pb[3] == 0:
                ok = False
                break
            ratios.append(pa[3] / pb[3])
        if not ok:
            continue
        # cluster ratios: find groups of >=3 with max/min-1 < 1e-6
        idx_sorted = sorted(range(len(ratios)), key=lambda k: ratios[k])
        i = 0
        while i < len(idx_sorted):
            j = i
            base = ratios[idx_sorted[i]]
            while j + 1 < len(idx_sorted) and abs(ratios[idx_sorted[j+1]] / base - 1) < 1e-6:
                j += 1
            grp = idx_sorted[i:j+1]
            if len(grp) >= 3:
                vals = [pairs[k][0][3] for k in grp]
                if interesting(vals):
                    scope = "SAME-FILE" if A["file"] == B["file"] else "CROSS-FILE"
                    w(f"[{scope} constant ratio ~{base:.9g} hits={len(grp)}/{len(pairs)}]")
                    w(f"  A: {label(A)}")
                    w(f"  B: {label(B)}")
                    w(f"  pairs: {[(pairs[k][0][0], pairs[k][0][3], pairs[k][1][0], pairs[k][1][3], f'{ratios[k]:.9g}') for k in grp]}")
                    cnt5 += 1
            i = j + 1
w(f"total ratio hits: {cnt5}\n")

# ---------- Detector 6: arithmetic / geometric progressions within a sequence ----------
w("##### DETECTOR 6: exact arithmetic/geometric runs (len>=4 consecutive) #####")
cnt6 = 0
for S in seqs:
    vals = [x[3] for x in S["items"]]
    coords = [x[0] for x in S["items"]]
    n = len(vals)
    # arithmetic
    for i in range(n):
        for L in range(4, n - i + 1):
            seg = vals[i:i+L]
            d = seg[1] - seg[0]
            if all(abs(seg[k+1] - seg[k] - d) < 1e-9 * max(1, abs(d)) for k in range(L-1)):
                # maximal check
                if i > 0 and abs(vals[i] - vals[i-1] - d) < 1e-9 * max(1, abs(d)):
                    continue
                if i + L < n and abs(vals[i+L] - vals[i+L-1] - d) < 1e-9 * max(1, abs(d)):
                    continue
                w(f"[arithmetic run len={L} d={d}] {label(S)} coords {coords[i:i+L]} vals {seg}")
                cnt6 += 1
                break
    # geometric
    for i in range(n):
        for L in range(4, n - i + 1):
            seg = vals[i:i+L]
            if any(v == 0 for v in seg):
                continue
            r = seg[1] / seg[0]
            if abs(r - 1) < 1e-12:
                continue
            if all(abs(seg[k+1] / seg[k] / r - 1) < 1e-9 for k in range(L-1)):
                if i > 0 and vals[i-1] != 0 and abs(vals[i] / vals[i-1] / r - 1) < 1e-9:
                    continue
                if i + L < n and abs(vals[i+L] / vals[i+L-1] / r - 1) < 1e-9:
                    continue
                w(f"[geometric run len={L} r={r:.9g}] {label(S)} coords {coords[i:i+L]} vals {seg}")
                cnt6 += 1
                break
w(f"total progression hits: {cnt6}\n")

# ---------- Detector 7: repeated fractional tails within a group ----------
w("##### DETECTOR 7: >=3 values in a group (n<=12) sharing identical decimal tail (>=3 decimals) #####")
cnt7 = 0
for S in seqs:
    if not (2 <= len(S["items"]) <= 12):
        continue
    tails = defaultdict(list)
    for coord, r, c, v in S["items"]:
        s = repr(v)
        if "." in s:
            frac = s.split(".")[1]
            if len(frac) >= 3:
                tails[frac].append((coord, v))
    for frac, lst in tails.items():
        if len(lst) >= 3:
            w(f"[shared tail .{frac} x{len(lst)}] {label(S)} -> {lst}")
            cnt7 += 1
# also long shared suffixes (>=6 trailing digits) between ANY two values in same sheet
w("--- DETECTOR 7b: pairs sharing >=6-digit decimal suffix within same sheet ---")
def dec_str(v):
    s = repr(v)
    return s.split(".")[1] if "." in s else ""
for (fn, sn), lst in by_sheet.items():
    items = [(coord, v) for r, c, coord, v in lst]
    for i in range(len(items)):
        for j in range(i+1, len(items)):
            a, b = dec_str(items[i][1]), dec_str(items[j][1])
            if len(a) >= 6 and len(b) >= 6:
                # longest common suffix
                k = 0
                while k < min(len(a), len(b)) and a[-1-k] == b[-1-k]:
                    k += 1
                if k >= 6:
                    w(f"[shared {k}-digit suffix] {fn} [{sn}] {items[i][0]}={items[i][1]!r} vs {items[j][0]}={items[j][1]!r} suffix=...{a[-k:]}")
                    cnt7 += 1
w(f"total tail hits: {cnt7}\n")

# ---------- Detector 8: exact duplicates within a group ----------
w("##### DETECTOR 8: exact duplicate values within one group (n<=12) #####")
cnt8 = 0
for S in seqs:
    if not (2 <= len(S["items"]) <= 12):
        continue
    seen = defaultdict(list)
    for coord, r, c, v in S["items"]:
        seen[round(v, 6)].append((coord, v))
    for val, lst in seen.items():
        if len(lst) >= 2 and interesting([val]):
            w(f"[dup {val} x{len(lst)}] {label(S)} -> {lst}")
            cnt8 += 1
w(f"total dup hits: {cnt8}\n")

# ---------- Detector 9: precision heterogeneity ----------
w("##### DETECTOR 9: decimal-precision distribution per sequence #####")
for S in seqs:
    if len(S["items"]) < 3:
        continue
    dist = defaultdict(int)
    for coord, r, c, v in S["items"]:
        dist[sig_decimals(v)] += 1
    if len(dist) >= 2:
        w(f'{label(S)}: decimals {dict(sorted(dist.items()))}')
w()

out.close()
print("done", pairs_reported, cnt4, cnt5, cnt6, cnt7, cnt8)
