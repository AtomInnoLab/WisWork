# PPT-P0-08 跨公司可比性：Apple 与 Toyota 候选材料包

状态：**官方财务原件、口径、复算、来源计划、八页参考稿和本机生产链路已验证；财务及使用权人工审阅、真实 PowerPoint 执行尚未完成**。本包计入候选材料，不计入 20 项通过数。固定资料截至 2024-11-01。

## 来源、报告期与比较边界

- Apple FY2024 截至 2024-09-28，审计 [Form 10-K](https://www.sec.gov/Archives/edgar/data/320193/000032019324000123/aapl-20240928.htm)；净销售额 **391,035 百万美元**，本地 PDF 第 32 页。仓库保存[Apple 投资者关系备案页](https://investor.apple.com/sec-filings/sec-filings-details/default.aspx?FilingId=17933082)直链的官方 PDF，并与 SEC 备案核对。
- Toyota FY2024 截至 2024-03-31，官方 [Form 20-F](https://global.toyota/pages/global_toyota/ir/library/sec/20-F_202403_final.pdf) 的本地 PDF 第 167 页列出合并销售收入 **45,095,325 百万日元**，其中金融服务收入 3,447,195 百万日元；会计准则为 IFRS。官方[年度财务汇总](https://global.toyota/pages/global_toyota/ir/financial-results/2024_4q_summary_en.pdf)本地 PDF 第 28 页在 **FY2024 已发生的 12 个月栏**列出 1 美元 = **145 日元**。右侧 FY2025 预测栏恰好也写 145，不作为本案例的历史汇率来源。
- 示意换算仅为 `45,095,325 / 145 = 311,002.24` **百万美元**，采用 Toyota FY2024 年度栏汇率；它不是 2024-11-01 即期汇率，也不代表 Apple 财年的平均汇率。不可据此与 Apple 净销售额相减、做排名或声称同口径规模。
- 财年结束日、US GAAP 与 IFRS、产品和金融服务收入范围均不同。`basis.json` 将同口径美元差额和跨公司收入排名明确设为 `null`。Toyota 20-F 于 2024-06-25、财务汇总于 2024-05-08、Apple 10-K 于 2024-11-01 披露，均不晚于固定截至日期；仍需财务审阅人确认主张与适用边界。

`basis.json` 与 `independent-recalc.csv` 冻结数值、公式、来源页和两个留空项。`reference-plan.json` 为五条带财务审阅状态的主张绑定七条快照来源，附两张图的原币源数值；`reference-deck.json` 保存逐页原生对象。`p0-08-reference.pptx` 是按该计划生成的八页原生参考稿：第 3 页用原生表格列出口径差异，第 7 页的不可比结果为真正**空白单元格**；Apple 和 Toyota 原币数值位于两张独立的原生图表，纵轴各自从零开始，没有混合币种或期间的比较图。运行 `node --import tsx generate-reference.cjs` 可重建，`node --import tsx verify-materials.mjs` 与 `sha256sum -c SHA256SUMS` 可核对来源快照、原件页码、字面数值、受限算术复算、图表底层值和留空项。

本机 `node tools/ppt-agent-electron-real-relay-smoke.mjs --benchmark=P0-08` 已通过真实 Electron PC、Rust Relay、Chrome 任务窗格、三份并发文档、来源上传、八页生产与 PPTX/PDF 回读，以及待处理和运行中任务恢复。浏览器侧仍是模拟 Office API 的任务窗格，不能替代真实 PowerPoint 保存关闭重开或金融结论验收。

## 固定任务提示

> 以本包 Apple 10-K、Toyota 20-F 和 Toyota 官方汇率汇总制作八页中文 16:9 原生可编辑汇报。先列公司、财年结束日、会计准则、币种、单位和收入范围，再给出可比与不可比矩阵。保留各自原币数值；只在独立区域展示 Toyota 按 FY2024 历史 12 个月汇率的示意美元换算，写明来源页、期间和公式。不得用该换算制造同口径差额或排名；证据不足的比较项留空并说明原因。保存来源、主张、逐页截图、图表底层值，以及 PowerPoint 保存关闭重开后的编辑证据。

## 历史获取记录与未完事项

旧版曾尝试 Apple 与 Sony 组合；Sony 官方 PDF 在本机网络路径返回 403。[获取状态记录](acquisition-status.json)保留该尝试，但 **Sony 不再是本包当前输入**。当前 Apple/Toyota 原件及 SHA256 已冻结。仍需人工核对两公司财务事实、期间适用性和公开文件使用条件，并执行真实 PowerPoint 八页任务及保存关闭重开检查；当前校验器只能证明材料与参考稿内部一致，不能给出投资判断或任务通过结论。
