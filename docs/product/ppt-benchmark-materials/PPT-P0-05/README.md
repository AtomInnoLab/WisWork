# PPT-P0-05：美国最高法院版权裁判比较材料候选

分析截止日固定为 **2024-12-01**，本次从官方站取得文件日期为 **2026-09-29**。这是以既有裁判作历史分析的材料包，不声称捕获了2024年的网站字节，也不声称裁判在其他时间或法域具有未经审阅的适用性。

## 原始资料

| 文件                                         | 案号 / 裁判日期     | 官方程序状态来源                                                                                                               |
| -------------------------------------------- | ------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| [Google 原件](google-oracle-official.pdf)    | 18-956 / 2021-04-05 | [官方 docket](https://www.supremecourt.gov/docket/docketfiles/html/public/18-956.html)：撤销并发回；2021-05-07 judgment issued |
| [Warhol 原件](warhol-goldsmith-official.pdf) | 21-869 / 2023-05-18 | [官方 docket](https://www.supremecourt.gov/docket/docketfiles/html/public/21-869.html)：维持；2023-06-20 judgment issued       |

下载 URL、完整案名、法院、实字节大小与 SHA-256 在 [manifest.json](manifest.json)；独立摘要见 [SHA256SUMS](SHA256SUMS)。两份均为官方发布的 slip opinion，保留正式出版前可修订提示；未冒称为最终装订版。原件包括 syllabus、正文及其他意见/附录，不删页、不改写。syllabus 是 Reporter 摘要，不属于法院意见。

## 八页独立范围与比较边界

[basis.json](basis.json) 给出恰好8页范围、逐项事实/争点/法院判断/明确限定/研究推论分类、比较矩阵及原始位置。PDF物理页从1计数，正文印刷页另计；引用不混用。Google 的假定可版权性、具体接口使用情境，与 Warhol 的特定商业许可及未评价用途必须保留。异议、下级法院意见不能写作多数意见。比较分析须经法律审阅，不把两个裁判组合成对第三个案的法律建议。

`node --import tsx generate-reference.cjs` 从冻结的两份官方 PDF 生成 [来源绑定计划](reference-plan.json)、[原生页模型](reference-deck.json) 与 [八页可编辑候选稿](p0-05-reference.pptx)。计划中的 12 个逐字摘录均取自多数意见正文，绑定原件 SHA-256 和物理页；7 项声明仍为 `needs_review`，比较声明标为研究判断。第 6 页用原生文本制作对象、用途、审查范围和结果矩阵，不嵌入裁判原件中的照片或艺术图像。候选稿是工程产物，不表示中文法律归纳或具体案件适用已经审阅。

本机 `node tools/ppt-agent-electron-real-relay-smoke.mjs --benchmark=P0-05` 已通过：两份冻结 PDF 上传、计划保存、八页生产、PPTX/PDF 回读和中断恢复。LibreOffice 已把候选 PPTX 渲染为八页 PDF 并目视检查，未见文字裁切或页面覆盖。这些验证未使用真实 PowerPoint 宿主，也不替代法律、隐私或图像权利审阅。

正文核对入口：[Google 官方PDF](https://www.supremecourt.gov/opinions/20pdf/18-956_new_0e04.pdf) 物理页5、15–19、26–32、39–40；[Warhol 官方PDF](https://www.supremecourt.gov/opinions/22pdf/21-869_87ad.pdf) 物理页7–19、27、29、43–44。原始PDF均完整保留，范围页并非删节原件。

## 公开原件与人工审阅

保留政府公开发布的裁判原件及原有公开姓名；**未执行额外脱敏，也未声称完成授权或隐私审阅**。Warhol 原件内含照片、艺术及杂志图像，官方发布不等于本材料已取得这些图像的独立复用许可；候选幻灯片仅做文字比较，不截取图像。法律、隐私、嵌入图像权利、布局与实际 PowerPoint 任务检查均待人工完成。工程生产不构成领域认证、宿主验收或任务 passed。

## 字节与定位校验

从仓库根运行：

```sh
node --import tsx docs/product/ppt-benchmark-materials/PPT-P0-05/verify-materials.mjs
```

需要 Node 与 Poppler (`pdfinfo`、`pdftotext`)。校验器检查 PDF 头、字节、SHA、真实页数、案号、8 页范围、声明引用是否处于所标正文页、产品解析器逐字摘录、计划与页模型绑定及 PPTX 原生页面；不认证专业结论、版权权限或布局。已对 Google 物理页 5 与 Warhol 物理页 27 做临时图像渲染检查，确认页号/多数意见标识及限定位置。图像检查不替代全部页面人工审阅。
