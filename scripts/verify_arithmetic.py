# Targeted verification of headline hits with exact arithmetic.
import json, os
from collections import defaultdict

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
cells = []
with open(os.path.join(BASE, "_dump", "cells.jsonl"), encoding="ascii") as f:
    for line in f:
        cells.append(json.loads(line))

grid = {}
for rec in cells:
    v = rec["value"]
    if isinstance(v, str):
        try: v = float(v)
        except Exception: continue
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        grid[(rec["file"], rec["sheet"], rec["row"], rec["col"])] = float(v)

def col(fn, sn, c, r0, r1):
    return [grid.get((fn, sn, r, c)) for r in range(r0, r1+1)]

out = open(os.path.join(BASE, "_dump", "verify.txt"), "w", encoding="utf-8")
def w(s=""): out.write(str(s) + "\n")

w("== 1. Fig1c vs Fig1f (MOESM7), rows 3-26, cols B and C ==")
b1 = col("MOESM7.xlsx","Fig.1c",2,3,26); b2 = col("MOESM7.xlsx","Fig.1f",2,3,26)
c1 = col("MOESM7.xlsx","Fig.1c",3,3,26); c2 = col("MOESM7.xlsx","Fig.1f",3,3,26)
w(f"C col identical positions: {sum(1 for a,b in zip(c1,c2) if a==b)}/24")
diffB = [(i+3, a, b) for i,(a,b) in enumerate(zip(b1,b2)) if a!=b]
w(f"B col identical positions: {24-len(diffB)}/24; differing rows (row,1c,1f): {diffB}")

w()
w("== 2. Fig1c col C vs MOESM10 ED1c col C (rounded) ==")
e1c = col("MOESM10.xlsx","Extended Data Fig. 1c",3,3,17)
for r,(a,b) in enumerate(zip(c1[:15], e1c), start=3):
    match = "MATCH" if abs(a-b) < 5e-6 else ""
    w(f"row{r}: Fig1c={a}  ED1c={b}  {match}")

w()
w("== 3. Fig1c (MOESM7) vs Fig3d (MOESM9): all column-pair exact ratios ==")
d = {}
for cF in range(2,7):
    vF = col("MOESM7.xlsx","Fig.1c",cF,3,26)
    for cG in range(2,13):
        vG = col("MOESM9.xlsx","Fig.3d",cG,3,26)
        pairs = [(a,b) for a,b in zip(vF,vG) if a is not None and b is not None and b!=0]
        if len(pairs) < 3: continue
        ratios = [a/b for a,b in pairs]
        # longest run of consecutive rows with constant ratio
        run_best = (0,None,None)
        i = 0
        while i < len(ratios):
            j = i
            while j+1 < len(ratios) and abs(ratios[j+1]/ratios[i]-1) < 1e-6:
                j += 1
            if j-i+1 > run_best[0]:
                run_best = (j-i+1, i, ratios[i])
            i = j+1
        if run_best[0] >= 3:
            from openpyxl.utils import get_column_letter as gcl
            w(f"Fig1c col {gcl(cF)} vs Fig3d col {gcl(cG)}: const ratio {run_best[2]:.10f} over rows {run_best[1]+3}..{run_best[1]+run_best[0]+2} (n={run_best[0]})")
            from fractions import Fraction
            w(f"   ratio as fraction: {Fraction(run_best[2]).limit_denominator(10000)}")

w()
w("== 4. MOESM10 ED1g/h/i/j col B vs MOESM12 ED4b/c/d/e col B ==")
for sm, s12 in [("Extended Data Fig. 1g","Extended Data Fig. 4b"),("Extended Data Fig. 1h","Extended Data Fig. 4c"),
                ("Extended Data Fig. 1i","Extended Data Fig. 4d"),("Extended Data Fig. 1j","Extended Data Fig. 4e")]:
    a = col("MOESM10.xlsx",sm,2,3,22); b = col("MOESM12.xlsx",s12,2,3,22)
    n_eq = sum(1 for x,y in zip(a,b) if x==y)
    w(f"{sm} vs {s12}: identical {n_eq}/20")

w()
w("== 5. MOESM11 ED3b vs MOESM15 ED10l rows3-10 cols B,C,D ==")
from fractions import Fraction
for c in (2,3,4):
    a = col("MOESM11.xlsx","Extended Data Fig. 3b",c,3,10)
    b = col("MOESM15.xlsx","Extended Data Fig. 10l",c,3,10)
    ratios = [x/y for x,y in zip(a,b)]
    ok = all(abs(ratios[0]/r-1) < 1e-9 for r in ratios)
    w(f"col {c}: constant ratio {ok}, ratio={ratios[0]:.15f}, fraction={Fraction(ratios[0]).limit_denominator(100000)}")

w()
w("== 6. MOESM15 ED10l internal ==")
for cols, name in [((2,5,6),"B/E/F"), ((4,8),"D/H")]:
    series = [col("MOESM15.xlsx","Extended Data Fig. 10l",c,3,12) for c in cols]
    for r in range(10):
        vals = [s[r] for s in series]
        w(f"row{r+3} {name}: {vals} {'ALL-EQUAL' if len(set(vals))==1 else ''}")
cC = col("MOESM15.xlsx","Extended Data Fig. 10l",3,3,12)
cG = col("MOESM15.xlsx","Extended Data Fig. 10l",7,3,12)
w("C vs G diffs: " + str([round(a-b,4) for a,b in zip(cC,cG)]))
for r in (11,12):
    row = [grid.get(("MOESM15.xlsx","Extended Data Fig. 10l",r,c)) for c in range(2,9)]
    ratios = [row[i+1]/row[i] for i in range(6)]
    w(f"row {r}: consecutive ratios {[f'{x:.12f}' for x in ratios]}")
b = col("MOESM15.xlsx","Extended Data Fig. 10l",2,3,12)
w(f"B9={b[6]} B11={b[8]} equal: {b[6]==b[8]}; B10={b[7]} B12={b[9]} equal: {b[7]==b[9]}")

w()
w("== 7. MOESM14 ED9f ==")
row7 = [grid.get(("MOESM14.xlsx","Extended Data Fig. 9f",7,c)) for c in range(2,14)]
w(f"row7 (replicate 5) B..M: {row7}")
for r_top, r_bot in ((3,8),(4,9),(5,10)):
    ratios = []
    for c in range(2,14):
        a = grid.get(("MOESM14.xlsx","Extended Data Fig. 9f",r_top,c))
        b = grid.get(("MOESM14.xlsx","Extended Data Fig. 9f",r_bot,c))
        ratios.append(None if b in (None,0) else a/b)
    hits = [(c+1, f"{r:.10f}") for c,r in enumerate(ratios) if r and abs(r/1.1764705882352942-1) < 1e-9]
    w(f"row{r_top}/row{r_bot}: ratio==20/17 at columns {[h[0] for h in hits]} ({len(hits)}/12)")
row6 = [grid.get(("MOESM14.xlsx","Extended Data Fig. 9f",6,c)) for c in range(2,8)]
w(f"row6 B..G: {row6}")
k = [grid.get(("MOESM14.xlsx","Extended Data Fig. 9f",r,11)) for r in range(3,11)]
w(f"col K rows3-10: {k}; diffs: {[k[i+1]-k[i] for i in range(len(k)-1)]}")

w()
w("== 8. qPCR 14-digit master-string clusters ==")
clusters = {
 "ED1a B3:B5 (MOESM10)": col("MOESM10.xlsx","Extended Data Fig. 1a",2,3,5),
 "ED7b B3:B5 (MOESM13)": col("MOESM13.xlsx","Extended Data Fig. 7b",2,3,5),
 "ED7b C3:C5": col("MOESM13.xlsx","Extended Data Fig. 7b",3,3,5),
 "ED7b B12:B14": col("MOESM13.xlsx","Extended Data Fig. 7b",2,12,14),
 "ED7b C12:C14": col("MOESM13.xlsx","Extended Data Fig. 7b",3,12,14),
 "ED9c B3:B5 (MOESM14)": col("MOESM14.xlsx","Extended Data Fig. 9c",2,3,5),
 "ED9c C3:C5": col("MOESM14.xlsx","Extended Data Fig. 9c",3,3,5),
 "ED9c D3:D5": col("MOESM14.xlsx","Extended Data Fig. 9c",4,3,5),
 "ED9c E3:E5": col("MOESM14.xlsx","Extended Data Fig. 9c",5,3,5),
 "ED1d B3:B5 (MOESM10)": col("MOESM10.xlsx","Extended Data Fig. 1d",2,3,5),
 "ED1d C3:C5": col("MOESM10.xlsx","Extended Data Fig. 1d",3,3,5),
 "ED1d D3:D5": col("MOESM10.xlsx","Extended Data Fig. 1d",4,3,5),
 "Fig3e B3:B5 (MOESM9)": col("MOESM9.xlsx","Fig.3e",2,3,5),
 "Fig3e C3:C5": col("MOESM9.xlsx","Fig.3e",3,3,5),
 "Fig3e L4:L5": col("MOESM9.xlsx","Fig.3e",12,4,5),
}
for k_, v in clusters.items():
    w(f"{k_}: {v}")
w()
w("exact diffs ED7b block2 - block1:")
for r in (3,4,5):
    dB = grid[("MOESM13.xlsx","Extended Data Fig. 7b",r+9,2)] - grid[("MOESM13.xlsx","Extended Data Fig. 7b",r,2)]
    dC = grid[("MOESM13.xlsx","Extended Data Fig. 7b",r+9,3)] - grid[("MOESM13.xlsx","Extended Data Fig. 7b",r,3)]
    w(f"row{r}: dB={dB!r} dC={dC!r}")
w()
w("== 9. MOESM11 ED3c OsRbohH block (rows 8-10) duplicates ==")
for c in range(2,8):
    vals = [grid.get(("MOESM11.xlsx","Extended Data Fig. 3c",r,c)) for r in (8,9,10)]
    w(f"col {c}: {vals}")
w()
w("== 10. MOESM11 ED3f selected duplicates ==")
for coord in [("N5",14,5),("AU5",47,5),("N6",14,6),("AU6",47,6),("P6",16,6),("AG6",33,6),("AB6",28,6),("AQ6",43,6),("BA6",53,6)]:
    name, c, r = coord
    w(f"{name} = {grid.get(('MOESM11.xlsx','Extended Data Fig. 3f',r,c))}")
out.close()
print("done")
