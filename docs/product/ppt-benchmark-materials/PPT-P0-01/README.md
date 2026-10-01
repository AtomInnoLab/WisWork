# PPT-P0-01 候选真实材料包

状态：**公开原始材料与八页可编辑候选参考稿已冻结，领域审阅与正式任务执行未完成**。本目录不计入 20 项专业任务通过数。

## 本机候选参考稿

`basis.json` 固定论文 PDF 第 5 页的 14/12 人样本、清单均分 1.6/6→2.2/6、`p=0.318`，以及 Table 1 “使用开源软件”7/14→10/12；第 9 页固定作者所述局限。`reference-plan.json` 是产品结构化制作计划，四条事实均标为 `needs_review`，来源摘录逐字匹配本机 PDF 解析器的原文；`reference-deck.json` 是与计划匹配的八页可编辑卡片。`node --import tsx generate-reference.cjs` 使用产品 PptxGenJS 编译器生成 `p0-01-reference.pptx`：八页原生文字和一张带嵌入 Excel 工作簿的原生图表，图表横轴分别写明两次样本分母。`node --import tsx verify-materials.mjs` 核对原 PDF 页数与摘录、计划/卡片绑定、图表缓存与工作簿值、八页结构和 SHA256SUMS。

本机 `node tools/ppt-agent-electron-real-relay-smoke.mjs` 已通过真实 Electron PC + Rust Relay 链路：上传冻结论文 PDF、保存计划、逐页生产八页，回读 PPTX/PDF，检查页面文字、原生图表及工作簿；该脚本同时验证并发与生产恢复。参考稿和本机链路仍是待科研审阅的候选内容，真实 PowerPoint 导入、编辑、保存、重开及专业任务验收尚未执行。

## 来源、版本和使用条件

- 论文：Ariel Deardorff, “Assessing the impact of introductory programming workshops on the computational reproducibility of biomedical workflows,” _PLOS ONE_ 15(7), e0230697，2020-07-08。[DOI](https://doi.org/10.1371/journal.pone.0230697)、[出版页面](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)、[论文 PDF](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0230697&type=printable)。本地 `deardorff-2020-article.pdf`，11 页，382769 字节，SHA256 `ff64709ef7b48a3292e4d4528809b50ed7f5d10836971b9cebcf35e769cf3617`。
- 补充材料：同一论文的 S1 Checklist，[出版页面](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)列出的 DOI `10.1371/journal.pone.0230697.s001`，[PDF](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0230697.s001&type=supplementary)。本地 `deardorff-2020-checklist.pdf`，1 页，47851 字节，SHA256 `c1bffef3f6b97236b2329b952791aa1e0d3176e2c7c5cd3e21a1e61e6e068ad9`。
- 论文出版页和 PDF 标示为 [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)；须保留作者、题名、来源、许可链接和修改说明。本目录保存原文件，未改写。获取日期：2026-09-28。

## 待人工核对的关键主张预标注

| ID  | 候选表述，须由科研审阅人确认                                                           | 原文位置                                                         | 核对重点                                                                           |
| --- | -------------------------------------------------------------------------------------- | ---------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| C1  | 研究在编程培训前访谈了 14 名 UCSF 生物医学研究人员，三个月后有 12 人完成后访谈。       | 论文 PDF 第 1、3、5 页，Methods、Results                         | 不得把两次样本量都写成 14，或把结果推广至所有生物医学研究者。                      |
| C2  | 六项清单平均分从培训前 1.6/6 到培训后三个月 2.2/6；论文报告差异未达统计显著，p=0.318。 | 论文 PDF 第 5 页，Checklist scores before and after the workshop | 均值变化不能写成已证明培训有因果效果；说明样本流失与统计功效。                     |
| C3  | Table 1 中“使用开源软件”为前测 7/14、后测 10/12；表格其余行为也需按各自分母展示。      | 论文 PDF 第 5 页，Table 1                                        | 图表须使用原生可编辑数据，并明确两列分母不同，不能直接把计数差解释为同一人的转变。 |
| C4  | 作者把小样本、招募/应答偏差、单人编码和定量分析功效不足列为局限。                      | 论文 PDF 第 9 页，Limitations                                    | 汇报的结论页必须同时呈现这些限制。                                                 |

以上为依据出版文本做的**预标注**，尚无领域审阅签名。S1 Checklist 的六项问题与 Table 1 对应，但它是空白测量表，不含参与者原始分数。若任务需要独立复算个体统计量，应另选获授权的底层数据并重新冻结材料版本。

## 固定任务提示

> 仅使用随任务提供的这篇论文 PDF 与 S1 Checklist，制作 8 页、16:9、中文为主的科研会议汇报。说明研究问题、样本与方法、六项清单、主要量化结果、定性观察和局限。至少制作一张引用 Table 1 的原生可编辑图表，清楚标注前测与后测不同的样本量；保留统计显著性结论，不能把相关观察写成因果证明。每条关键主张标出论文页码和原文位置；无法由材料支持的内容明确标为假设或待核验。逐页检查内容、结构、几何、视觉和来源，交付来源/QA 摘要。最终文件须在 PowerPoint 保存、关闭、重开，并演示修改标题与图表数据后再保存。

## 执行前仍需完成

1. 科研审阅人核对上述预标注、图表数值与引用位置，签署日期；必要时修订固定提示并提升材料包版本。
2. 在获准的真实 PowerPoint 环境记录 PC/Addin/宿主版本、材料哈希、文档及项目身份、时间线、逐页截图和原生图表回读。
3. 使用 P0 验收清单保存 PPTX、重开编辑录屏及其 SHA256，并记录所有首次失败和后续尝试。当前未执行。
