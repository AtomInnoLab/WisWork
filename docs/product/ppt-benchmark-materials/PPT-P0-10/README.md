# PPT-P0-10 大型真实 PDF 工程验证材料

来源：[IPCC Special Report on the Ocean and Cryosphere in a Changing Climate 完整报告](https://www.ipcc.ch/site/assets/uploads/sites/3/2022/03/SROCC_FullReport_FINAL.pdf)。本机于 2026-09-28 下载并核对：**51,870,607 字节（49.47 MiB）**、766 页、SHA256 `cadeefed4b0f0627384b6b7f3730afc729570270b8794b71759f9dc6511a36b2`。生产 PDF 解析器提取 4,338,831 个 UTF-16 字符、766 个页面段落，其中末页无可提取文本。本机原件位于 `/tmp/wiswork-p0-10-srocc-full-report.pdf`，仓库不保存全文；重新下载后须先核对字节数和 SHA256。使用与再分发边界见 [IPCC 版权说明](https://www.ipcc.ch/copyright/)。

运行真实原件回归：`WISWORK_P0_10_PDF=/tmp/wiswork-p0-10-srocc-full-report.pdf node_modules/.bin/vitest run apps/shell/tests/presentation-attachments-integration.test.ts -t 'real near-limit PDF'`。此测试从 Office 插件客户端分块发送给 PC 服务，完成生产解析、附件元数据验证，并读取第 4,000,000 字符处的 2,000 字符窗口。2026-09-28 实测通过；此前 100 万字符边界下在 `attachment_finish` 失败。单次读取窗口保持 24,000 字符上限，网页快照仍保持 100 万字符上限。

## 固定 8 页摘要范围草案

1. 报告范围、年份与证据级别。
2. 冰冻圈与高山地区变化（第 2 章）。
3. 极地地区变化（第 3 章）。
4. 海平面上升的观测与投影（第 4 章）。
5. 海洋与海洋生态系统变化（第 5 章）。
6. 极端事件及风险（第 6 章）。
7. 跨章节关键图表、数值和不确定性对照。
8. 适应选项、限制及原文来源索引。

此范围只供专业审阅者确认，不能视为原文结论。正式 P0-10 尚需在 PowerPoint 宿主中记录上传和解析耗时、制作并复核 8 页摘要、核对原文表格与结论、保存重开并验证原生编辑；当前不计入已通过任务。
