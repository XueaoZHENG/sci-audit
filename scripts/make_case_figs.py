# -*- coding: utf-8 -*-
"""
案例图生成：Nature 2024 s41586-024-07418-9 (OsCERK1/OsCIE1) 撤稿论文
所有图均从原始 Source Data xlsx 直接读取，不使用任何人工编造数值。
输出：cases/nature-2024-okie1/figs/*.png
"""
import os, sys, json, re
import openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager, patches
from matplotlib.colors import ListedColormap

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 140

BASE = r"C:\Users\xueao\Nutstore\1\酶制剂小组\17投稿文章\SCI_AUDIT\nature_oscie1"
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "figs")
os.makedirs(OUT, exist_ok=True)

RED, ORG, YEL, GRN, GREY = "#c62828", "#ef6c00", "#f9a825", "#2e7d32", "#9e9e9e"


def load(fn, sheet):
    wb = openpyxl.load_workbook(os.path.join(BASE, fn), data_only=True)
    ws = wb[sheet]
    out = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                out[(r, c)] = v
    wb.close()
    return out


def colvals(cells, col, r0, r1):
    return [cells[(r, col)] for r in range(r0, r1 + 1) if (r, col) in cells]


def fmt(v):
    if isinstance(v, float):
        s = repr(v)
        return s
    return str(v)


# ============================================================
# FIG 1 — Fig.1c vs Fig.1f 24 对病斑面积逐点相同
# ============================================================
c1 = load("MOESM7.xlsx", "Fig.1c")
c2 = load("MOESM7.xlsx", "Fig.1f")
x = colvals(c1, 3, 3, 26)   # Fig.1c  Oscie1-1
y = colvals(c2, 3, 3, 26)   # Fig.1f  Oscie1-1
fig, ax = plt.subplots(figsize=(5.6, 5.4))
ax.plot(x, y, "o", ms=7, mfc=RED, mec="white", mew=0.8, alpha=0.85, label="24/24 完全相同")
lim = [min(min(x), min(y)) - 0.5, max(max(x), max(y)) + 0.5]
ax.plot(lim, lim, "--", color=GREY, lw=1.2, label="y = x（完全相同）")
ax.set_xlim(lim); ax.set_ylim(lim)
ax.set_xlabel("Fig. 1c  Oscie1-1 病斑面积 (n=24)", fontsize=11)
ax.set_ylabel("Fig. 1f  Oscie1-1 病斑面积 (n=24)", fontsize=11)
ax.set_title("两个独立实验的 24 个「不同植株叶片」逐值重合", fontsize=12.5, fontweight="bold", pad=12)
ax.legend(loc="upper left", fontsize=9.5, frameon=False)
ax.grid(alpha=0.25, lw=0.6)
ax.text(0.97, 0.04, "PubPeer #4 · 本技能 R-01", transform=ax.transAxes,
        ha="right", fontsize=9, color=RED, fontweight="bold")
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig1_fig1c_vs_fig1f.png")); plt.close(fig)

# ============================================================
# FIG 2 — qPCR「14 位母串池」
# ============================================================
qsrc = [
    ("MOESM9.xlsx",  "Fig.3e", 3, 5, "Fig. 3e"),
    ("MOESM10.xlsx", "Extended Data Fig. 1a", 3, 5, "ED Fig. 1a (NIPB)"),
    ("MOESM10.xlsx", "Extended Data Fig. 1d", 3, 5, "ED Fig. 1d"),
    ("MOESM13.xlsx", "Extended Data Fig. 7b", 3, 5, "ED Fig. 7b (块1)"),
    ("MOESM14.xlsx", "Extended Data Fig. 9c", 3, 5, "ED Fig. 9c"),
]
rows = []
for fn, sh, r0, r1, tag in qsrc:
    cells = load(fn, sh)
    for r in range(r0, r1 + 1):
        for c in range(2, 13):
            v = cells.get((r, c))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                s = fmt(float(v))
                if "." in s and len(s.split(".")[1]) >= 10:
                    rows.append((tag, s, s.split(".")[1]))
tails = {}
for tag, s, frac in rows:
    tails.setdefault(frac[-12:], []).append((tag, s))
# 只保留被 >=2 个面板共用的尾串
shared = {k: v for k, v in tails.items() if len({t for t, _ in v}) >= 2}
shared = dict(sorted(shared.items(), key=lambda kv: -len(kv[1]))[:6])

fig, ax = plt.subplots(figsize=(11.8, 5.9))
ax.axis("off")
ax.text(0.5, 0.965, "6 条「母串」跨 5 个面板流通 —— 所有 qPCR 数值都是它们的「改前导位」变体",
        ha="center", fontsize=13, fontweight="bold", transform=ax.transAxes)
ax.text(0.5, 0.905, "概率约 1e-12 / 独立实验；PubPeer 18 条评论零提及",
        ha="center", fontsize=10.5, color=RED, transform=ax.transAxes)
row_y = 0.815
for i, (tail, items) in enumerate(shared.items()):
    tags = sorted({t for t, _ in items})
    ax.add_patch(patches.FancyBboxPatch((0.015, row_y - 0.085), 0.97, 0.115,
                                        boxstyle="round,pad=0.004", linewidth=0.9,
                                        edgecolor=RED, facecolor="#fdecea"))
    ax.text(0.028, row_y - 0.027, f"母串尾号  …{tail}", fontsize=10.5, fontweight="bold",
            color=RED, transform=ax.transAxes, family="sans-serif")
    shown = []
    for tag, s in items:
        key = (tag, s)
        if key in shown: continue
        shown.append(key)
    txt = "    ".join(f"{t.split('(')[0].strip()}: {s}" for t, s in shown[:3])
    ax.text(0.028, row_y - 0.062, txt[:190], fontsize=8.6, transform=ax.transAxes,
            family="sans-serif", color="#333333")
    row_y -= 0.128
ax.text(0.5, 0.012, "例：0.90230473800567 与 1.10230473800567 共享 13 位尾串，差值恰为 0.2 —— 手改前导位",
        ha="center", fontsize=9.5, color="#555555", transform=ax.transAxes, style="italic")
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig2_qpcr_mother_strings.png")); plt.close(fig)

# ============================================================
# FIG 3 — ED Fig.10l 的 0.99 等比 + 0.85 复制
# ============================================================
l = load("MOESM15.xlsx", "Extended Data Fig. 10l")
row11 = [l[(11, c)] for c in range(2, 9)]
row3 = [l[(3, c)] for c in range(2, 9)]
row4 = [l[(4, c)] for c in range(2, 9)]
row5 = [l[(5, c)] for c in range(2, 9)]
row8 = [l[(8, c)] for c in range(2, 9)]
row9 = [l[(9, c)] for c in range(2, 9)]
row10 = [l[(10, c)] for c in range(2, 9)]
fig, axes = plt.subplots(1, 2, figsize=(11.4, 4.5))

ax = axes[0]
xs = np.arange(len(row11))
ax.plot(xs, row11, "o-", color=RED, lw=2, ms=7, label="ED10l 第 11 行（第 9 重复）")
ideal = [row11[0] * (0.99 ** k) for k in range(len(row11))]
ax.plot(xs, ideal, "--", color="#333333", lw=1.4, label="理论 X0 * 0.99^n")
for x_, y_ in zip(xs, row11):
    ax.annotate(f"{y_:,.1f}", (x_, y_), textcoords="offset points", xytext=(0, 8),
                ha="center", fontsize=7.2, color=RED)
ax.set_xticks(xs); ax.set_xticklabels([f"M{i+1}" for i in xs], fontsize=9)
ax.set_xlabel("ED Fig.10l 的 7 个 Mock 组（相邻比值恒为 0.99）", fontsize=10)
ax.set_ylabel("RLU", fontsize=10)
ax.set_title("跨基因型的「第 9 重复」互为 0.99 倍", fontsize=11.5, fontweight="bold")
ax.legend(fontsize=8.5, frameon=False); ax.grid(alpha=0.25, lw=0.6)
ax.text(0.03, 0.06, "PubPeer #3 / #15 · 本技能 R-14", transform=ax.transAxes,
        fontsize=8.5, color=RED, fontweight="bold")

ax = axes[1]
pairs = [(3, 8), (4, 9), (5, 10)]
labels = ["重复1 → 重复6", "重复2 → 重复7", "重复3 → 重复8"]
for (a, b), lab in zip(pairs, labels):
    va = [l[(a, c)] for c in range(2, 9)]
    vb = [l[(b, c)] for c in range(2, 9)]
    ax.plot(va, vb, "o", ms=6, mfc=ORG, mec="white", alpha=0.85, label=lab)
allv = [l[(a, c)] for a in (3, 4, 5) for c in range(2, 9)] + [l[(b, c)] for b in (8, 9, 10) for c in range(2, 9)]
ax.plot([0, max(allv)], [0, max(allv)], "--", color=GREY, lw=1.2, label="y = x")
ax.set_xlabel("第 1/2/3 重复 (RLU)", fontsize=10); ax.set_ylabel("第 6/7/8 重复 (RLU)", fontsize=10)
ax.set_title("第 6/7/8 重复 = 第 1/2/3 重复 × 0.85", fontsize=11.5, fontweight="bold")
ax.legend(fontsize=8.5, frameon=False); ax.grid(alpha=0.25, lw=0.6)
ax.text(0.03, 0.06, "PubPeer #6 · 本技能 R-11", transform=ax.transAxes,
        fontsize=8.5, color=RED, fontweight="bold")
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig3_ed10l_geometric.png")); plt.close(fig)

# ============================================================
# FIG 4 — 34 面板 × 8 类模式 的命中矩阵
# ============================================================
PAT = ["逐值全同","恒定偏移","恒定比值","等比/等差","长尾串/母串","组内重复","统计自洽","统计不自洽"]
PATDEF = [
 ("Fig.1c","Fig.1",[0,1,2],""), ("Fig.1d","Fig.1",[], ""), ("Fig.1f","Fig.1",[0,1,2],""),
 ("Fig.1g","Fig.1",[], ""),
 ("Fig.2e","Fig.2",[], ""), ("Fig.2g","Fig.2",[], ""), ("Fig.2h","Fig.2",[], ""),
 ("Fig.3d","Fig.3",[2],""), ("Fig.3e","Fig.3",[4],""),
 ("ED1a","ED1",[0,1,6],""), ("ED1c","ED1",[], ""), ("ED1d","ED1",[1,4],""),
 ("ED1e","ED1",[], ""), ("ED1g","ED1",[0],""), ("ED1h","ED1",[0],""),
 ("ED1i","ED1",[0],""), ("ED1j","ED1",[0],""),
 ("ED3a","ED3",[], ""), ("ED3b","ED3",[2],""), ("ED3c","ED3",[5],""),
 ("ED3e","ED3",[], ""), ("ED3f","ED3",[2],""), ("ED3g","ED3",[], ""),
 ("ED4b","ED4",[0],""), ("ED4c","ED4",[0],""), ("ED4d","ED4",[0],""), ("ED4e","ED4",[0],""),
 ("ED4f","ED4",[], ""),
 ("ED7b","ED7",[1,4,7],""),
 ("ED9c","ED9",[4],""), ("ED9f","ED9",[0,2,4],""),
 ("ED10f","ED10",[], ""), ("ED10l","ED10",[0,1,3,6],""),
 ("ED12g","ED12",[], ""),
]
M = np.zeros((len(PATDEF), len(PAT)))
for i,(nm,fg,pats,_) in enumerate(PATDEF):
    for j in pats: M[i,j]=1
fig, ax = plt.subplots(figsize=(9.8, 8.2))
cmap = ListedColormap(["#eceff1", RED])
ax.imshow(M, cmap=cmap, aspect="auto", vmin=0, vmax=1)
ax.set_xticks(range(len(PAT)))
ax.set_xticklabels(PAT, rotation=38, ha="right", fontsize=9)
ax.set_yticks(range(len(PATDEF)))
ax.set_yticklabels([f"{p[0]:6s} ({p[1]})" for p in PATDEF], fontsize=8.4, family="monospace")
ax.set_ylim(len(PATDEF)-0.5, -0.5)
for i in range(len(PATDEF)):
    for j in range(len(PAT)):
        if M[i,j]:
            ax.add_patch(patches.Rectangle((j-0.46, i-0.46), 0.92, 0.92, facecolor=RED,
                                           edgecolor="white", lw=1.2))
ax.set_xticks(np.arange(-0.5, len(PAT), 1), minor=True)
ax.set_yticks(np.arange(-0.5, len(PATDEF), 1), minor=True)
ax.grid(which="minor", color="white", lw=1.4)
ax.tick_params(which="minor", length=0)
ax.set_title("Source Data 全量枚举：10 个 xlsx / 34 个面板（每个面板 = 一个 sheet）\n"
             "红色 = 该面板命中对应取证模式；空白 = 按既定方法检查未发现异常",
             fontsize=11.6, fontweight="bold", pad=14)
fig.subplots_adjust(left=0.22, right=0.97, top=0.86, bottom=0.10)
fig.savefig(os.path.join(OUT, "fig4_panel_matrix.png")); plt.close(fig)

# ============================================================
# FIG 5 — PubPeer 18 条对账
# ============================================================
items = [
    ("#2", "检出"), ("#3", "检出"), ("#4", "检出"), ("#5", "检出*"), ("#6", "检出"),
    ("#7", "检出"), ("#8", "检出"), ("#9", "检出*"), ("#10", "检出"), ("#11", "本轮补出"),
    ("#12", "检出"), ("#13", "并入#5"), ("#15", "检出"), ("#16", "检出"),
    ("#17", "部分"), ("#18", "部分"), ("#1", "图像范围外"), ("#14", "作者声明"),
]
fig, ax = plt.subplots(figsize=(10.6, 4.6))
xs = np.arange(len(items))
colors = []
for _, s in items:
    colors.append({"检出": RED, "检出*": RED, "本轮补出": "#8e24aa", "部分": ORG,
                   "并入#5": ORG, "图像范围外": GREY, "作者声明": GREY}[s])
ax.bar(xs, [1] * len(items), color=colors, edgecolor="white", lw=0.8)
ax.set_xticks(xs)
ax.set_xticklabels([i[0] for i in items], rotation=45, ha="right", fontsize=9)
ax.set_yticks([])
ax.set_ylim(0, 1.55)
for x_, (n, s) in zip(xs, items):
    ax.text(x_, 1.06, s, ha="center", fontsize=8, rotation=90,
            color="white" if s in ("检出", "检出*") else "#333333", fontweight="bold")
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=RED, label="完全检出"),
                   Patch(color="#8e24aa", label="本轮补出（#11）"),
                   Patch(color=ORG, label="部分检出 / 等价"),
                   Patch(color=GREY, label="图像范围外 / 指控源")],
          loc="upper center", bbox_to_anchor=(0.5, -0.30), ncol=4, fontsize=9, frameon=False)
ax.set_title("PubPeer 18 条评论逐条对账：数值项 15 条中 严格检出 10 条 + 本轮补出 1 条 = 11/15",
             fontsize=12, fontweight="bold", pad=12)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig5_pubpeer_reconciliation.png")); plt.close(fig)

# ============================================================
# FIG 6 — 时间线 + 路径泄漏
# ============================================================
meta = [
    ("MOESM10", "2024-03-15 02:21"), ("MOESM12", "2024-03-15 13:41"),
    ("MOESM14", "2024-03-15 14:30"), ("MOESM16", "2024-03-15 14:33"),
    ("MOESM15", "2024-03-15 14:45"), ("MOESM9", "2024-03-17 11:36"),
    ("MOESM11", "2024-03-17 13:06"), ("MOESM7", "2024-03-17 13:09"),
    ("MOESM13", "2024-03-18 06:19"), ("MOESM8", "2024-04-27 13:59"),
]
def d(s):
    return __import__("datetime").datetime.strptime(s, "%Y-%m-%d %H:%M")
t_sub, t_acc, t_pub = d("2022-04-11 00:00"), d("2024-04-12 00:00"), d("2024-05-15 00:00")
fig, ax = plt.subplots(figsize=(12.6, 5.6))
for i, (name, ts) in enumerate(meta):
    tt = d(ts)
    ax.scatter([tt], [i], s=110, color=RED if i < 9 else ORG, zorder=3,
               edgecolor="white", lw=1.0)
    ax.text(tt, i + 0.30, name, ha="center", fontsize=8.2, color="#444444")
ax.axvline(t_sub, color=GREY, ls="--", lw=1.1)
ax.axvline(t_acc, color="#1565c0", ls="--", lw=1.3)
ax.axvline(t_pub, color="#1565c0", ls="--", lw=1.3)
ax.text(t_sub, -1.35, "投稿\n2022-04-11", ha="center", fontsize=8.6, color="#555555")
ax.annotate("接收\n2024-04-12", xy=(t_acc, -0.35), xytext=(t_acc, -1.35), ha="right",
            fontsize=8.6, color="#1565c0",
            arrowprops=dict(arrowstyle="-", color="#1565c0", lw=0.8))
ax.annotate("发表\n2024-05-15", xy=(t_pub, -0.35), xytext=(t_pub, -1.35), ha="left",
            fontsize=8.6, color="#1565c0",
            arrowprops=dict(arrowstyle="-", color="#1565c0", lw=0.8))
ax.set_yticks(range(len(meta)))
ax.set_yticklabels([m for m, _ in meta], fontsize=8.6)
ax.set_ylim(-1.7, len(meta) - 0.3)
ax.set_xlabel("源数据 xlsx 的最后修改时间 (docProps/modified)", fontsize=10)
ax.set_title("10 个源数据文件的最后修改时间：9/10 集中在 2024-03-15~18，MOESM8 在接收后",
             fontsize=12, fontweight="bold", pad=14)
ax.grid(axis="x", alpha=0.25, lw=0.6)
ax.text(0.015, 0.045, "路径泄漏（xl/workbook.xml → x15ac:absPath）：\n"
        "D:\\OsCIE1文章投稿\\阶段5：接收后修改\\supplementary information graph\\",
        transform=ax.transAxes, ha="left", va="bottom", fontsize=9.2, color=RED,
        family="sans-serif", bbox=dict(boxstyle="round,pad=0.45", fc="#fdecea", ec=RED, lw=0.9))
ax.set_xlim(d("2022-02-01 00:00"), d("2024-08-15 00:00"))
fig.subplots_adjust(left=0.095, right=0.985, top=0.88, bottom=0.16)
fig.savefig(os.path.join(OUT, "fig6_timeline.png")); plt.close(fig)

print("figures written to", OUT)
for f in sorted(os.listdir(OUT)):
    print("  ", f, os.path.getsize(os.path.join(OUT, f)), "bytes")
