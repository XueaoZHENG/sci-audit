# scripts/ 说明

本目录是 `sci-audit` 技能的**参考实现**。全部脚本都来自 `cases/nature-2024-okie1` 案例的真实审查过程（已对该撤稿论文跑通并与手算复核一致）。

| 脚本 | 作用 | 对应 SKILL.md |
|---|---|---|
| **`sci_audit_probe.py`** | **推荐入口**。一条命令跑完：全量转储 → T1 数值指纹(P1–P8) → T2 低样本(L1–L5) → T0 元数据/容器取证 → 统计量自洽性复算。输出全部落文件，不进上下文 | §〇 执行预算、§三 全部阶段 |
| `detect_patterns.py` | 通用数值指纹探针：逐值全同 / 恒定偏移 / 恒定比值 / 等差等比 / 长尾串 / 组内重复 / 精度分布 | §三 3.1 |
| `probe_lown.py` | 低样本 n≤3 专项（补通用阈值的盲区）。技术要点：直接解析 `xl/worksheets/sheetN.xml` 取 `<v>` 原始字面量，用最短往返 repr 计算有效小数位 | §三 3.2 |
| `probe_meta.py` | 元数据取证：docProps、`x15ac:absPath` 绝对路径、zip 成员时间戳、公式 vs 手填列级清单、统计量重算 | §三 5.1 |
| `stats_consistency.py` | 统计量自洽性复算（Mean Diff. / t）。**注意：必须配 `--control` 已知干净面板** | §三 3.4 |
| `verify_arithmetic.py` | 命中项的精确算术复核 | — |
| `make_case_figs.py` | 从 Source Data 重绘案例图（散点/母串池/等比数列/面板矩阵/对账条形图/时间线） | `cases/` |
| `annotate_pubpeer.py` | 给 PubPeer 原图叠加「判定横幅 + 检出机制 + 报告章节」 | `cases/` |

## 用法

```bash
# 主入口
python sci_audit_probe.py --input <含 xlsx 的目录> --out <输出目录> --control <已知干净面板名关键词>
# 产出：dump/cells.jsonl、findings_numeric.txt、findings_lown.txt、
#       findings_meta.txt、findings_stats.txt、summary.md

# 单独跑某一项（脚本内 BASE 常量指向案例数据目录，按需改）
python probe_lown.py
python stats_consistency.py
```

## 环境

```bash
pip install numpy scipy openpyxl pillow matplotlib
```

## 已知限制（用前必读）

1. **`sci_audit_probe.py` 的 t 检验配对**：当同一 sheet 内有多个 t 块、且候选数据列 > 2 时，脚本**不硬判**，改为列出候选列并标注"需人工指定配对"（GraphPad 一个块可含多组两两比较，自动配对极易错）。案例中的 ED7b 判定即为人工指定配对后手算复核所得。
2. **Excel 解析**：脚本直接读 `xl/worksheets/sheetN.xml` 与 `xl/sharedStrings.xml`。
   - 字符串单元格是 sharedStrings 的**索引**，必须解析，否则表头全是数字；
   - 自闭合的 `<c r="A2" s="2"/>` 与 `<c r="B2" ...>...</c>` 必须区分，否则会**静默吞掉下一格**。
3. **统计量比对要按存储显示精度归一**：GraphPad 存的是四舍五入后的显示值，拿全精度重算值直接比必然出伪 delta。
4. **误报控制**：块分类 `DATA / SEQ / STAT` 必须先做（人工序号列与 GraphPad 派生统计量会产生上百万条噪声）。
