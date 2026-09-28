# PPT-P0-03 候选真实材料包

状态：**论文、数据字典及去个体化汇总表已准备，领域审阅与正式任务执行未完成**。本目录不计入 20 项专业任务通过数。

## 来源与许可

- George Hess、M. Nils Peterson, “Bicycles May Use Full Lane” Signage Communicates U.S. Roadway Rules and Increases Perception of Safety, _PLOS ONE_ 10(8), e0136973，2015-08-28。[出版页面](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0136973)、[论文 PDF](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0136973&type=printable)。本地 `hess-peterson-2015-article.pdf`，16 页。
- 出版方附录 [S1 Data 数据字典](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0136973.s001&type=supplementary)，本地 `hess-peterson-2015-dictionary.pdf`，2 页。[S2 Data 原始 CSV](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0136973.s002&type=supplementary)，原文件 SHA256 `c73b2ac2f1e2aacefb0d568ac5650464c6dcabb9fd1a75733475db8534e44136`；因含逐条时间戳与自由文本，本目录只保留可复算汇总 `treatment-outcomes.csv`，不随仓库分发原始行。
- 出版页声明 [CC BY](https://creativecommons.org/licenses/by/4.0/)。保留作者、题名、出版来源、许可链接，并标明汇总表是本项目从 S2 Data 转换生成。获取日期：2026-09-28。目录内文件哈希见 `SHA256SUMS`。

## 汇总表口径与待审预标注

S2 Data 含 1,824 行美国受访者。以 `Treatment` 为组；`n` 为该组行数；四个 `*_agree` 列是对应字典字段严格等于 `1_Agree` 的行数。分组和为 489+422+454+459=1,824。用 Python 标准库可按以下逻辑重算：

```python
import csv
from collections import Counter, defaultdict

groups = defaultdict(list)
with open('hess-peterson-2015-data.csv', newline='') as source:
    for row in csv.DictReader(source):
        groups[row['Treatment']].append(row)
for treatment, rows in sorted(groups.items()):
    counts = Counter(
        (field, row[field])
        for row in rows
        for field in ('Permitted2', 'Safe2', 'Permitted4', 'Safe4')
    )
    print(treatment, len(rows), [counts[field, '1_Agree'] for field in ('Permitted2', 'Safe2', 'Permitted4', 'Safe4')])
```

| ID  | 候选主张，须由交通研究审阅人确认                                         | 原文位置与检查点                                                                 |
| --- | ------------------------------------------------------------------------ | -------------------------------------------------------------------------------- |
| C1  | 论文收到 1,978 份问卷，排除 154 份非美国答卷，分析 1,824 份。            | 论文 PDF Results 与 Table 1；不得把总回复数当作分析分母。                        |
| C2  | 汇总表的四组样本量分别为 489、422、454、459；汇总值是回答“同意”的计数。  | S2 Data 的 `Treatment` 与四个结果字段，S1 Data 定义；百分比须以各组 `n` 为分母。 |
| C3  | 这项网络问卷比较道路标志的理解与主观安全感，不直接测量实际交通事故变化。 | 论文 Abstract、Methods、Study Limitations；不应表述为已证实降低事故。            |

上述是**预标注**，没有领域审阅签名。汇总表不包含作者的回归调整、置信区间或因果估计，不能用它宣称复现了论文的统计模型。论文讨论美国道路标志，展示时需说明地域适用范围与非随机招募限制。

## 固定任务提示

> 仅使用随任务提供的论文 PDF、数据字典 PDF 和汇总 CSV，制作 8 页、16:9、中文为主的交通研究汇报。说明研究问题、问卷设计、样本筛选、四组标志、两种道路场景、理解与安全感结果及研究局限。至少制作一张来自汇总 CSV 的原生可编辑图表，展示各组 `Permitted2` 和 `Safe2` 的同意比例，并同时写明各组分母和计算式。不要把主观安全感写成真实事故率；不要把未在材料中重算的回归结果写成已独立复现。逐页标明关键主张在论文中的位置或汇总字段；交付来源与 QA 摘要。最终文件须在 PowerPoint 保存、关闭、重开，并演示修改图表数据后再保存。

## 执行前仍需完成

1. 交通研究审阅人核对汇总、提示与关键主张，签署日期；必要时修订并冻结材料版本。
2. 在获准的真实 PowerPoint 环境记录 PC/Addin/宿主版本、材料哈希、文档及项目身份、时间线、逐页截图、原生图表回读，以及保存关闭重开证据。
3. 保存首次失败和后续尝试，按 P0 验收清单逐项评分。当前未执行。
