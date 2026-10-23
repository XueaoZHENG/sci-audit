# sci-audit — SCI 论文数据真实性审查 Skill

一个用于 AI 编程助手（Claude Code / Kimi Code 等）的审查技能：对**生物类 SCI 论文**做系统性的**造假与数据真实性审查**。

![案例：Nature 2024 撤稿论文的 18 张 PubPeer 原图标注卡](cases/nature-2024-okie1/figs/pubpeer/pubpeer_04_Fig.1c-Fig.1f.png)

## 覆盖范围

- **图像完整性**：凝胶 / Western blot / CoIP / pull-down 的条带复制、拼接、跨图复用检测，**内参跨图复用**这一生物论文头号雷区专项检查
- **生物学实验专项清单**：Western blot、琼脂糖凝胶、RT-qPCR、CFU 菌落计数、RNA-seq/蛋白组学、HPLC-MS、表型照片、Y2H/split-LUC、显微照片，逐类列出必查项
- **数值一致性**：源数据 Excel 与发表图逐项对账、统计复算（mean/SD/95CI/t 检验）、发表图散点像素级识别逐行比对
- **组学复核**：DEG/DEP 计数、Venn 集合运算、热图声称复现、**accession 真实存在且公开**
- **色谱/质谱轨迹**：复制粘贴铁证检测（非零完全相同片段、基线噪声独立性）
- **元数据取证**：Office docProps 时间线、**`x15ac:absPath` 泄漏的绝对路径**、公式派生 vs 手填值
- **数据可用性**：Figshare/PDB/BMRB 等声明链接的**实际可达性与内容比对**

## 特色

- **判定分级对齐 ORI/COPE**：🔴 红线（fabrication/falsification 铁证）/ 🟠 存疑 / 🟡 合规瑕疵（honest error）/ 🟢 通过（带覆盖率量化），不冤枉真实数据
- **执行预算分层（T0–T3）+ Token 纪律**：把 80% 的检出率用 20% 的 token 拿到；全量数据只进文件不进上下文
- **强制执行的专业报告模板**：审查范围矩阵（每个检查项一行）、红线清单、未覆盖范围声明
- **统计取证方法族**：GRIM / GRIMMER / SPRITE / statcheck 式 p 值复算 / 末位数字分析（附适用条件）
- **对照纪律**：统计量自洽性复算**必须配已知干净面板**，否则极易得出伪结论（见案例 §四）
- **内置假阳性清单**：凝胶边缘、marker 重复上样、CFU 离散计数重复、色谱共同峰、饱和条带、归一化锚点 1.0、P 值网格撞值、人工序号列、Tukey CI 半宽恒定
- **案例库**：[`cases/nature-2024-okie1/`](cases/nature-2024-okie1/) —— 在一篇已撤稿的 Nature 论文上盲审复现，量化检出率/特异性/token 成本

## 快速开始

```bash
# 1) 安装：把 SKILL.md 放进助手的技能目录
#    Claude Code : ~/.claude/skills/sci-audit/SKILL.md
#    Kimi Code   : ~/.kimi-code/skills/sci-audit/SKILL.md

# 2) 一条命令跑完数值 + 元数据取证（输出全部落文件，不进上下文）
python scripts/sci_audit_probe.py --input <含 xlsx 的目录> --out <输出目录> --control <已知干净面板名>

# 3) 然后对助手说
#    「审查这篇论文的数据真实性」
```

## 目录

```
SKILL.md                     技能主文件（判定分级 / 报告模板 / 执行预算 / 全流程）
README.md                    本文件
scripts/
  sci_audit_probe.py         合并探针：T0 元数据 + T1 数值指纹 + T2 低样本 + 统计自洽性
  detect_patterns.py         通用数值指纹探针（identical run / offset / ratio / progression / tail）
  probe_lown.py              低样本 n<=3 专项（首值相同 / 整数偏移 / 精度断层 / 粒度网格）
  probe_meta.py              元数据取证（docProps / absPath / zip 时间戳 / 公式 vs 手填）
  stats_consistency.py       统计量自洽性复算（Mean Diff. / t，含对照面板）
  make_case_figs.py          案例图生成（从 Source Data 重绘）
  annotate_pubpeer.py        PubPeer 原图标注卡生成
cases/
  nature-2024-okie1/         完整案例（图文并茂 + 对账表 + 12 条修订输入）
```

## 判定分级（对齐 ORI/COPE）

| 级别 | 定义 |
|---|---|
| 🔴 红线 | 有像素级/数值级**铁证**指向 fabrication 或 falsification |
| 🟠 存疑 | 现有材料无法定论，需要原始材料/作者澄清 |
| 🟡 合规瑕疵 | 属于 honest error 范畴，不影响结论可靠性 |
| 🟢 通过 | 按既定方法检查未发现异常（须注明覆盖率） |

## 来源

方法论蒸馏自一篇已发表 eLife 论文的真实全链条自查实践，并在一篇已撤稿 Nature 论文上做了外部盲审验证。

案例论文：Wang G. *et al.* Nature **629**, 1158–1164 (2024). DOI `10.1038/s41586-024-07418-9`
