# PPT-P0-17 三文档并行与另存隔离候选材料

状态：**三套来源、各自来源绑定的八页原生候选稿与故障步骤已冻结；科研、法律、财务审阅及真实三窗口 PowerPoint 验收未完成**。本包不计入 20 项任务通过数。

## 三套互不混淆的输入

| 文档                          | 来源、版本与范围                                                                                                                                                                                                                                                                                                                                                                              | 8 页目标识别码 |
| ----------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------- |
| `deardorff-2020-article.pdf`  | Deardorff，_PLOS ONE_ 15(7), e0230697 (2020)，DOI `10.1371/journal.pone.0230697`，11 页；[出版页](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)，CC BY 4.0；与 [P0-01](../PPT-P0-01/README.md) 原件摘要一致                                                                                                                                                      | `SCI-P0-17`    |
| `edpb-guidelines-07-2020.pdf` | 欧洲数据保护委员会 Guidelines 07/2020，Version 2.1，封面载明 2021-07-07 通过，版本史载明 2022-09-20 小幅修订，51 页；[官方发布页](https://www.edpb.europa.eu/documents/guideline/guidelines-072020-on-the-concepts-of-controller-and-processor-in-the-gdpr_en)和[官方 PDF](https://www.edpb.europa.eu/system/files/documents/2023-10/EDPB_guidelines_202007_controllerprocessor_final_en.pdf) | `LAW-P0-17`    |
| `apple-fy2024-form10k.pdf`    | Apple FY2024 Form 10-K，报告期截至 2024-09-28，121 页；与 [P0-07](../PPT-P0-07/README.md) 已冻结原件摘要一致；来源及 SEC 对照见 P0-07                                                                                                                                                                                                                                                         | `FIN-P0-17`    |

文件 SHA256 见 `SHA256SUMS`。三份 PDF 分别用于科研、官方规则说明和审计财务披露，只作为本机基准输入；不能把 EDPB 指南冒充法规原文，也不能由指南推导个案法律结论。不同材料版本、许可与专业判断由对应审阅人复核。冻结日期 2026-09-28。

## 固定任务与故障时点

`scenario.json` 固定三份资料、各自可见识别码、8 页目标标题和五步交错操作。三项目均要求中文为主、16:9、原生可编辑 PPTX。封面及交付摘要必须显示本项目识别码；页级计划和来源只允许引用本项目输入。法律稿只解释指南文本及其版本，不提供法律意见；财务稿只报告 FY2024 披露值，不虚构预测；科研稿保留方法和局限。

在同一 PC 登录下打开三个 PowerPoint 窗口：交错完成至少两页后关闭科研文档，法律文档另存为新文件，财务文档继续。随后向关闭前及另存前会话注入延迟响应，核对旧文档和其他项目均未被写入。每个项目继续至 8 页，分别保存、关闭、重开并编辑原生对象。保留三窗口录屏、各自 project/request/document/session/slide ID、旧会话拒绝记录、三份文件摘要、逐页截图与五层 QA。单项目成功不替代另两个项目。

`node --import tsx generate-reference.cjs` 从三个本地 PDF 分别构建 `reference-science-*`、`reference-legal-*`、`reference-finance-*` 的计划、页模型与原生 PPTX。三个计划均只绑定本项目 PDF 的 SHA-256 和逐页原文摘录；每页保留识别码，科研稿与财务稿第 6 页包含来源绑定原生图表。`node --import tsx verify-materials.mjs` 核对 13 份冻结文件、PDF 页数和来源标识、24 页计划与编译后 PPTX，并拒绝跨项目来源及识别码混淆。校验器不运行 Office.js，也不能证明真实窗口隔离、迟到响应被拒绝或资料结论准确。

本机生产链路可分别运行 `node tools/ppt-agent-electron-real-relay-smoke.mjs --benchmark=P0-17:science --built-taskpane`，将 `science` 改为 `legal` 或 `finance` 可验证另两份。三个变体各上传自己的 PDF；本机整链路只验证独立生产，交错三窗口、另存及迟到响应仍须专门验证。

## 正式验收仍需完成

1. 科研、法律、财务审阅人分别核对来源和任务边界，登记姓名与日期；EDPB 文档的效力、适用时点及 Apple 财务口径不能由本包自动判断。
2. 用目标 Windows/Mac/Web PowerPoint 真实运行三个窗口及另存/关闭/迟到响应矩阵；录屏并导出诊断、回执、文档身份和三份独立交付物。
3. 按 [P0 验收清单](../../wiswork-ppt-agent-p0-acceptance-2026-09-23.md) 保存每份 PPTX 的重开编辑证据，并核查没有跨文档写入、漏页、重复页或虚假完成。
