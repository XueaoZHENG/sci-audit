# -*- coding: utf-8 -*-
"""
把 PubPeer 评论的原图标注成本技能的"检出卡片"：
  顶部：判定横幅（评论号 / 面板 / 检出状态）
  底部：检出机制 + 报告章节 + 是否超出 PubPeer 自身发现
原图保持不动（PubPeer 评论者的红框原样保留），只加横幅与脚注。
输出：cases/nature-2024-okie1/figs/pubpeer/*.png
"""
import os, sys, glob, re
from collections import defaultdict
from PIL import Image, ImageDraw, ImageFont

SRC = r"C:\Users\xueao\Nutstore\1\酶制剂小组\17投稿文章\SCI_AUDIT\nature_oscie1\_dump\pubpeer_imgs"
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    r"C:\Users\xueao\Nutstore\1\酶制剂小组\17投稿文章\SCI_AUDIT\nature_oscie1", "figs", "pubpeer")
os.makedirs(OUT, exist_ok=True)

F_REG = r"C:\Windows\Fonts\msyh.ttc"
F_BLD = r"C:\Windows\Fonts\msyhbd.ttc"
RED, ORG, PUR, GREY, INK = "#c62828", "#ef6c00", "#6a1b9a", "#616161", "#1a1a1a"

# 评论号 -> (面板, 判定, 徽章色, 机制, 报告章节, 是否超越 PubPeer 本身)
INFO = {
 1:  ("ED Fig. 6a", "本轮范围外", GREY, "图像复制（Y2H 斑点）—— 本轮按用户指示未做图像完整性审查",
      "§6 未覆盖范围", "—"),
 2:  ("ED Fig. 10l", "已检出", RED, "跨基因型共享完全相同的 8 连值块 + ±200000 / ±80000 恒定整数偏移",
      "R-13", "一致"),
 3:  ("ED Fig. 10l", "已检出", RED, "第 9/10 行为精确等比数列，公比恒为 0.99（X0*0.99^n）",
      "R-14", "一致"),
 4:  ("Fig. 1c / Fig. 1f", "已检出", RED, "Oscie1-1 组 24/24 逐值全同；NIPB 组 12/24 相同 + 恒定整数偏移",
      "R-01 / R-02", "一致"),
 5:  ("Fig. 1c", "已检出", RED, "行 18-26：Oscie1-1 / Oscie1-2 比值恒定 0.980392157（9 对）",
      "R-18", "一致（检测器已命中，上一轮正文漏收）"),
 6:  ("ED Fig. 9f", "已检出", RED, "第 6/7/8 重复 = 第 1/2/3 重复 x 0.85（17/20），27/27 对命中",
      "R-11", "一致"),
 7:  ("ED Fig. 9f", "已检出", RED, "组内重复尾数家族；K 列精确等差 +300000",
      "R-12", "一致 + 超出（K 列等差为本次新增）"),
 8:  ("ED Fig. 1a / 7b", "已检出", RED, "两图版 14 位小数逐值全同（跨文件）",
      "R-21", "一致"),
 9:  ("ED Fig. 3f", "已检出", RED, "spl11-1 / spl11-2 样本 7 三对比值恒定 0.967105",
      "R-17", "一致（上一轮已命中但正文误归因）"),
 10: ("ED Fig. 3c", "已检出", RED, "OsRbohH 三基因型组内各有 2 个值 14 位小数级完全相同",
      "R-16", "一致"),
 11: ("ED Fig. 1a", "本轮补出", PUR, "两个 OE 株系首值逐位相同(11.2) + 2/3 对恒定 +1.0 + 13 位精度断层",
      "R-05", "一致（上一轮漏检，本轮由低样本专项补出）"),
 12: ("Fig. 1c / Fig. 1f", "已检出", RED, "NIPB 组逐行恒定整数偏移 +2.0/-4/-1/+3/-5/+6",
      "R-02", "一致"),
 13: ("Fig. 1c", "已检出", RED, "比值 0.980392157（#5 的图片复述）",
      "R-18", "并入 #5"),
 15: ("ED Fig. 10l", "已检出", RED, "第 9/10 行 6 个基因型互为 0.99 倍；行 9/行 10 恒定 1.3844501535",
      "R-14", "一致（本轮多覆盖 1 列）"),
 16: ("ED Fig. 3b / 10l", "已检出", RED, "两图版 Mock 组精确恒定比值 99/86 = 1.151162790（24/24 对）",
      "R-09", "一致"),
 17: ("Fig. 1c / Fig. 3d", "部分检出", ORG, "精确有理数比例链 41/38（COMWT-1 9/9 + Oscie1-1 3/9）；NIPB 8 对因舍入未达阈值",
      "R-03", "一致"),
 18: ("Fig. 1f / Fig. 3d", "部分检出", ORG, "精确有理数比例链 41/38（Oscie1-1 3/9）；NIPB 4 对因舍入未达阈值",
      "R-03", "一致"),
 14: ("作者声明（通讯作者 Ertao Wang）", "指控源", GREY,
      "点名 Fig.1c 与 ED Fig.1a/3c/4f/7b/9f/10l 的数值问题，及 Fig.1b / ED Fig.6a 的图像复制；"
      "一作数据造假属实，论文主动撤稿。声明所列 7 个数值面板，本技能覆盖 5 个完全检出、"
      "1 个由低样本专项补出、1 个（ED Fig.4f）专项复核后证据不足",
      "§2 红线清单 + §4.1 存疑", "指控源，不计入灵敏度分母"),
}


def font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def wrap(draw, text, fnt, maxw):
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur); cur = ""; continue
        t = cur + ch
        if draw.textlength(t, font=fnt) > maxw:
            lines.append(cur); cur = ch
        else:
            cur = t
    if cur: lines.append(cur)
    return lines


files = sorted(glob.glob(os.path.join(SRC, "*.png")))
bycomment = defaultdict(list)
for f in files:
    m = re.match(r"comment(\d+)_", os.path.basename(f))
    if m:
        bycomment[int(m.group(1))].append(f)

made = []
for n in sorted(set(list(bycomment) + [14])):
    if n not in INFO:
        continue
    panel, verdict, badge, mech, ref, extra = INFO[n]

    imgs = [Image.open(f).convert("RGB") for f in bycomment.get(n, [])]
    if not imgs:
        # 无附图的评论（如作者声明）→ 渲染一张纯文字卡
        import json as _json
        from pathlib import Path as _Path
        raw = _Path(SRC).parent / "pubpeer_comments_raw.json"
        body = ""
        try:
            cs = _json.load(open(raw, encoding="utf-8"))
            body = (cs[n - 1].get("markdown") or "")
        except Exception:
            pass
        body = re.sub(r"!\[[^\]]*\]\([^)]*\)", "[附图]", body)
        body = body.replace("**", "").replace("`", "")
        im = Image.new("RGB", (1200, 60), "white")
        dd = ImageDraw.Draw(im)
        ff = font(F_REG, 17)
        yy = 16
        for ln in wrap(dd, body, ff, 1160)[:14]:
            dd.text((20, yy), ln, font=ff, fill="#333333")
            yy += 24
        im = im.crop((0, 0, 1200, max(yy + 12, 60)))
        imgs = [im]
    W = max(im.width for im in imgs)
    H = sum(im.height for im in imgs) + 8 * (len(imgs) - 1)
    BW = max(W, 980)
    BH = 96 + H + 128
    canvas = Image.new("RGB", (BW, BH), "white")
    d = ImageDraw.Draw(canvas)

    f_h1 = font(F_BLD, 26)
    f_h2 = font(F_REG, 18)
    f_foot = font(F_REG, 16)
    f_footb = font(F_BLD, 16)

    # ---- 顶部横幅 ----
    d.rectangle([0, 0, BW, 96], fill="#212121")
    d.rectangle([0, 0, 12, 96], fill=badge)
    d.text((32, 16), f"PubPeer #{n}", font=f_h1, fill="#ffffff")
    d.text((32, 56), f"{panel}", font=f_h2, fill="#bdbdbd")
    label = f"本技能判定：{verdict}"
    tw = d.textlength(label, font=f_h1)
    d.rounded_rectangle([BW - tw - 46, 30, BW - 26, 74], radius=8, fill=badge)
    d.text((BW - tw - 36, 38), label, font=f_h1, fill="#ffffff")

    # ---- 原图（多张时纵向堆叠）----
    yy = 96
    for im in imgs:
        canvas.paste(im, ((BW - im.width) // 2, yy))
        d.rectangle([(BW - im.width) // 2, yy, (BW - im.width) // 2 + im.width, yy + im.height],
                    outline="#e0e0e0", width=1)
        yy += im.height + 8

    # ---- 底部说明 ----
    y = 96 + H + 18
    d.rectangle([0, y, BW, BH], fill="#fafafa")
    d.line([(0, y), (BW, y)], fill="#e0e0e0", width=1)
    d.text((28, y + 14), "检出机制", font=f_footb, fill=INK)
    lines = wrap(d, mech, f_foot, BW - 180)
    d.text((130, y + 14), lines[0], font=f_foot, fill=INK)
    for k, ln in enumerate(lines[1:]):
        d.text((130, y + 14 + 22 * (k + 1)), ln, font=f_foot, fill=INK)
    y2 = y + 14 + 22 * len(lines) + 6
    d.text((130, y2), f"报告章节：{ref}", font=f_footb, fill=RED)
    d.text((130, y2 + 24), f"与 PubPeer 原文的关系：{extra}", font=f_foot, fill="#555555")
    d.text((130, y2 + 48), "sci-audit 技能 · 案例 nature-2024-okie1", font=f_foot, fill="#9e9e9e")

    out = os.path.join(OUT, f"pubpeer_{n:02d}_{panel.replace(' ', '').replace('/', '-')}.png")
    canvas.save(out)
    made.append((n, os.path.basename(out), canvas.size))

for n, name, size in sorted(made):
    print(f"#{n:2d}  {name}  {size}")
print(f"\n{len(made)} annotated cards -> {OUT}")
