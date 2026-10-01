# PPT-P0-10 — SROCC 八页摘要候选材料

状态：候选材料已冻结，**范围未经人工确认；专业、版权分发、人工版面与真实 PowerPoint 宿主验收均未完成**。这不是完整报告的替代读物，不意味着真实性、支持关系、权威性或时效性已通过。

## 原件身份及时间边界

IPCC 官方 [SROCC 全文](https://www.ipcc.ch/site/assets/uploads/sites/3/2022/03/SROCC_FullReport_FINAL.pdf)，766 个实际 PDF 页；本机外部路径 `/tmp/wiswork-p0-10-srocc-full-report.pdf`，可用环境变量 `WISWORK_P0_10_PDF` 显式指定同一原件。

- 字节数：51,870,607（约 49.47 MiB）。
- SHA-256：`cadeefed4b0f0627384b6b7f3730afc729570270b8794b71759f9dc6511a36b2`。
- 科学评估年份 2019；此全文版版权页为 2022 出版，不能理解为科学证据更新到 2022 或 2026。PDF 第 14 页脚注 3：文献提交截止 2018-10-15、接受截止 2019-05-15。
- 原全文不入 Git，且不生成替代或填充 PDF。读取附件时使用上述原件真实字节；各 source 的 attachment URI 对应其原件 SHA。

## 冻结的八页范围

以下为自制中文摘要；实际 PDF 页码从第一页算起，括号为原文印刷页。不是挑选八张原文页面复制。完整摘要、限定、来源链接与原件定位保留在 basis、plan、deck 和各页备注。

| 摘要页 | 范围                     | 原文定位                                       | 必须保留的限定                                                                    |
| ------ | ------------------------ | ---------------------------------------------- | --------------------------------------------------------------------------------- |
| 1      | 阅读范围、版本与不确定性 | PDF 9（ix）；14（4）引言/脚注 3、6             | 证据/一致性置信度不同于概率；年份不是最新有效性认证                               |
| 2      | 高山与冰冻圈观测口径     | PDF 16（6），A.1.1–A.1.3/脚注 9、10            | 全球冰川不等于高山样本；2006–2015 质量损失；格陵兰含外围冰川                      |
| 3      | 极地差异                 | PDF 16（6），A.1.4                             | 北极 1979–2018 九月海冰趋势不能代替南极；海冰不能代替冰盖                         |
| 4      | 海平面观测与投影         | PDF 20（10）A.3.1、A.3.4；30（20）B.3.1–B.3.2  | 2100 时点不同于 2081–2100 均值；RCP2.6/8.5 相对 1986–2005；保留可能范围与地方差异 |
| 5      | 海洋热、酸化与缺氧       | PDF 19（9）A.2.1–A.2.5；20（10）A.2.6          | 期间、深度与测量口径各自保留；不能合并为单个变化率                                |
| 6      | 极端事件                 | PDF 19（9）A.2.3；30（20）B.3.4；31（21）B.3.6 | 暴露/情景/置信度保留；部分地点不是所有地点，未来气旋频率仍低置信度                |
| 7      | 数值对照与算术           | PDF 16（6）A.1.1/脚注10；30（20）B.3.1         | Gt/yr、mm/yr、m 不互换；220/360 仅量纲算术，不复现不确定性或认证气候估计          |
| 8      | 适应、限制与全文索引     | PDF 40（30）C.2；41（31）C.3                   | 因地制宜、治理/生态/财务限制，措施不保证有效；必要时读完整章节                    |

完整章节执行摘要导航：高山 PDF143（133）、极地215（205）、海平面333（323）、海洋460（450）、极端事件601（591）。本候选没有声称已完整审阅这五章或其引用研究。

原文 `confidence`、`likelihood` 与可能范围不同；中文摘要保留这些边界。plan 的 low confidence/needs_review 表示本候选尚待审阅，不替代 IPCC 自己的置信度标签。数值期与上下界见 basis；禁止省掉原期间、单位、情景、基线及区间后作图。

## 可恢复材料与核验

- `basis.json`：冻结范围、原文位置、数值与未完成审阅项。
- `reference-plan.json`：严格生产 schema、逐页 claims/source 与完整 science context；无新增工具。
- `reference-deck.json`、`p0-10-reference.pptx`：八页原生可编辑文字和备注；不含原图、截图或原全文，不声称包含原生表格/图表。
- `generate-reference.cjs`：使用已有默认生产 compiler；`verify-materials.mjs`：核对外部原件身份、实际原文页锚点、plan/deck、八页原生文字、产物 hashes。校验通过不证明候选内容充分、范围获确认、专业审阅或宿主验收通过。
- `manifest.json`：原件外部身份及冻结产物哈希。

从仓库根运行：

```sh
node --import tsx docs/product/ppt-benchmark-materials/PPT-P0-10/verify-materials.mjs
node --import tsx docs/product/ppt-benchmark-materials/PPT-P0-10/generate-reference.cjs
WISWORK_P0_10_PDF=/tmp/wiswork-p0-10-srocc-full-report.pdf node_modules/.bin/vitest run apps/shell/tests/presentation-attachments-integration.test.ts -t 'real near-limit PDF'
```

重新生成会改变 PPTX 字节；须重新人工核查并冻结 manifest，verifier 不自动修改 hashes。原件若遗失，明确重新获取并核对，不自动网络重放：

```sh
curl --fail --location --output /tmp/wiswork-p0-10-srocc-full-report.pdf https://www.ipcc.ch/site/assets/uploads/sites/3/2022/03/SROCC_FullReport_FINAL.pdf
sha256sum /tmp/wiswork-p0-10-srocc-full-report.pdf
stat -c %s /tmp/wiswork-p0-10-srocc-full-report.pdf
pdfinfo /tmp/wiswork-p0-10-srocc-full-report.pdf
```

## 已取得的机器证据与限制

2026-10-01 本机补跑真实 Electron PC + Rust Relay + 构建版 Taskpane 业务冒烟：使用上述 **51,870,607 字节原 PDF**，由冻结清单核对 SHA256 后分块上传并由 PC 解析，保存计划、逐页生产八页，核对 PPTX/PDF 回读；同次公共链路的浏览器配对、三文档隔离、断线续会、待处理任务与运行中任务恢复均通过，脚本退出码 0。P0-10 来源任务段耗时 **115.627 秒**，不包含意图、Agent 研究、真实 PowerPoint 导入/QA 或专业审阅，不能用作用户完成时间或 P95。实测发现旧候选稿的状态行越过来源脚注安全区，现已调整并重新冻结；八页在 LibreOffice 转 PDF 后逐页目视检查，材料校验器新增编译器几何门禁。上述结果不证明原文摘要充分、版权分发获准、真实 PowerPoint 保存重开或专业任务通过。

2026-09-29 新鲜生产定向测试：**1 passed / 8 skipped**，总 30.83s（不是产品解析延迟承诺）。实际 PC 服务与 AttachmentSkill 上传/解析真实近 50 MiB PDF，读取第 4,000,000 字符之后，并核对摘录来源定位；日志外部 `/tmp/p0-10-production.log`。此前实际提取 4,338,831 UTF-16 字符，末页空白仍占一页；本次测试也确认大于四百万字符。未运行全套回归。

2026-09-30 逐页扫描图检测启用后的本机复验：同一 SHA256 原件再次经 Office AttachmentSkill 分块上传、真实 PC 服务解析、400 万字符之后的窗口读取和来源页定位，定向测试 **1 passed / 9 skipped**，总 50.04s，其中测试主体 46.43s。此结果证明该原件在当前逐页检测路径可完成，不是 50 MiB PDF 的统一性能保证，也不代替 Relay 网络及真实 PowerPoint 宿主验收。

原 PDF16、30 已用 Poppler 渲染并检查值、情景、基线、置信度位置；八页候选另由 LibreOffice 作临时 PDF 版面检查。LibreOffice 检查不等于真实 PowerPoint 重开、编辑、另存、导入验收。

## 版权与分发边界

已核对 IPCC [版权政策](https://www.ipcc.ch/copyright/) 官方检索结果及原 PDF 第2页版权页：在线版列 CC-BY-NC-ND 4.0，第三方材料可能另有权利；IPCC 政策对少量个人非商业引用要求完整出处。政策网页直接访问返回 403，因此不声称已取得完整当前网页正文或分发许可。本仓库仅保存少量自制摘要、定位与可编辑候选；不保存/再分发原全文、不复制图表、长段原文或第三方图片。本候选仍需版权及用途审阅，不能把引用说明当作许可批准。
