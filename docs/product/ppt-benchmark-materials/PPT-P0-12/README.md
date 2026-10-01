# PPT-P0-12 候选材料包：品牌模板与字体回退

状态：**真实专业来源、自制品牌模板和八页来源绑定候选已冻结；品牌/科研审阅及真实 PowerPoint 执行未完成**。本包不计入 20 项任务通过数。

`generate-candidate.mjs` 使用本包两份 PDF、Brand Kit 和 P0-01 的冻结科研候选内容生成 `reference-plan.json`、`reference-deck.json`；额外为 S1 Checklist 建立独立来源与待审主张。八页采用模板封面和内容槽位，原生图表保持 7/14、10/12 的不同分母。`verify-candidate.mjs` 核对来源摘要、完整来源映射、槽位几何、字体候选和图表数据。本机 Electron PC + Rust Relay + 构建版 Taskpane 冒烟上传两份 PDF、生产八页并回读 PPTX/PDF；真实字体探针确认缺失的 `WisWork Benchmark Display 2026` 回退到已安装的 `Noto Sans CJK SC`。八页已用 LibreOffice 渲染逐页检查；该结果不证明真实 PowerPoint 字形或专业内容通过。

## 材料与授权

| 文件                                    | 用途                                           | 来源与使用条件                                                                                                                                                                                                               |
| --------------------------------------- | ---------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `deardorff-2020-article.pdf`            | 专业内容原文，11 页                            | 与 [PPT-P0-01](../PPT-P0-01/README.md) 相同的 PLOS ONE 论文；[出版页](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)，[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)。原 PDF 未修改。 |
| `deardorff-2020-checklist.pdf`          | 研究量表，1 页                                 | 同论文 S1 Checklist；[出版页](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)，CC BY 4.0。原 PDF 未修改。                                                                                         |
| `wiswork-benchmark-brand-template.pptx` | 3 页可编辑品牌样式示例                         | WisWork 为本基准任务自制，无第三方品牌图形；由 `generate-template.mjs` 生成。它是模板样本，不是 8 页交付稿。                                                                                                                 |
| `brand-kit.json`                        | 可录入 Presentation Brand Kit 的颜色和布局槽位 | WisWork 自制版本 1；使用前按当前 PC 版本校验。                                                                                                                                                                               |

冻结日期：2026-09-28。文件 SHA256 见 `SHA256SUMS`。两份 PDF 是从 PPT-P0-01 包复制的独立冻结副本，内容与哈希一致；此任务另有品牌模板及字体回退约束，因此不能用 PPT-P0-01 的结果代替本任务结果。重新生成模板或修订品牌约束时须提升材料版本并重跑。

## 品牌约束与预期检查

- 16:9，底色白，墨蓝 `#102A43`、青绿 `#007F86` 为主；其他允许颜色仅见 `brand-kit.json`。标题、正文、证据侧栏和页脚的位置以模板及 Brand Kit 为参照，不可只复制背景色。
- 标题指定字体 `WisWork Benchmark Display 2026` 是**故意缺失**的测试字体；执行时记录宿主字体清单和实际替代字体，选一款宿主已安装、覆盖所需中文字形的字体，记录替代前后截图和行高/溢出检查。不得声称指定字体已安装或替代必然一致。
- 冻结计划的 `style.fontFace` 可保持上述指定字体，并设置 `style.fontFallbacks`（如本机已安装的 `Noto Sans CJK SC`）作为有序候选。PC 的 PptxGenJS 编译报告记录 `fontResolution.requested/used/substituted`；候选均不可用会返回 `font_unavailable`。本机可用性不证明 PowerPoint 最终字形、行宽或跨设备一致，仍须逐页截图审阅。
- 若走 Office.js 的 `add_slide_ir_objects` 直接写入路径，须把已检查的 `fontResolution.used` 或明确审阅的候选字体填入 `resolved_font_face`；工具只接受 `style.fontFace` 或 `style.fontFallbacks` 中的字体，缺少选择时写前拒绝。该字段是显式选择，不是 Office 宿主字体探针。
- 正文优先使用 `Noto Sans CJK SC`；若宿主缺失，同样记录实际回退及字形检查。页脚必须保留来源与日期；原生图表须标明分母、单位、期间和来源。
- 模板的 3 页分别展示封面、内容证据布局、图表布局。最终 8 页须根据专业内容填写，原生文本/形状/表格/图表可编辑；模板占位语句不能作为研究事实交付。

## 固定任务提示

> 仅使用本包两份科研 PDF 作为专业内容，以本包品牌 PPTX 和 Brand Kit 为约束，在 PowerPoint Taskpane 中制作 8 页、16:9、中文为主的科研会议汇报。覆盖研究问题、方法与样本、六项清单、量化结果、定性观察和局限；原生图表注明不同时间点的分母。关键主张逐条引用原文页码，不能把未显著的结果表述为已证实的效果。保持模板中的标题、证据区、颜色、页脚和图表约束。标题故意使用缺失字体，请在真实宿主中选择并记录可用回退字体，检查全部 8 页没有丢字、越界或风格漂移。交付 PPTX、Claim Ledger、五层 QA 摘要和逐页截图；保存、关闭、重开后修改标题与图表数据再保存。

## 执行前仍需完成

1. 品牌审阅人确认模板约束、替代字体判定和布局容差；科研审阅人核对 P0-01 原文预标注，签署日期。
2. 在目标 PowerPoint Windows/Mac/Web 宿主记录字体清单、实际回退、8 页截图、对象可编辑性与保存重开；记录模板与最终页面的视觉偏差。
3. 使用 [P0 验收清单](../../wiswork-ppt-agent-p0-acceptance-2026-09-23.md)归档全部材料哈希、任务身份、诊断和首次/补跑尝试。当前均未执行。
