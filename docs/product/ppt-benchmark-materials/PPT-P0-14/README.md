# PPT-P0-14：外部素材失败与受控降级

状态：**候选材料包与八页来源绑定 SlideIR 已冻结，待科学解释与使用权人工审阅；真实 PowerPoint 任务未执行**。本包满足固定专业主题资料、两个稳定官方图片 URL、许可线索、本地原图备份和一次受控超时脚本的准备要求；它不是 P0 任务通过记录。

## 固定材料与表达范围

- 主题为 **2024 年全球地表温度异常**，固定资料时点为 2025-01-11。两份原始官方网页已保存为 `nasa-2024-article.html` 和 `nasa-svs-5450.html`。NASA 的[地球观测文章](https://science.nasa.gov/earth/earth-observatory/2024-was-the-warmest-year-on-record-153806/)与[SVS 5450 图像页](https://svs.gsfc.nasa.gov/5450)均说明：2024 年全球平均温度比 NASA 的 1951–1980 年基线高 **1.28 °C**。这是异常值，不是绝对温度，也不能推出任意地区的精确温度。
- 官方[4K 原图](https://svs.gsfc.nasa.gov/vis/a000000/a005400/a005450/2024GISTEMPMap.png)为首选，[2K 原图](https://svs.gsfc.nasa.gov/vis/a000000/a005400/a005450/2024GISTEMPMap_2K.png)为降级候选。二者均有本地 PNG 备份、尺寸和 SHA256 记录。SVS 页面要求署名 NASA Scientific Visualization Studio 及 NASA/GSFC GISS 数据贡献者。[NASA 媒体使用说明](https://www.nasa.gov/nasa-brand-center/images-and-media/)一般允许教育/信息用途，但标识、可识别人员、第三方素材及商业用法例外仍须人工核对；`asset-rights.json` 不构成最终法律许可结论。
- `p0-14-reference.pptx` 是八页 **PptxGenJS 工程参考稿**：文字和流程表格为 PowerPoint 原生对象，地图是嵌入的原始 PNG。已用 LibreOffice 渲染八页并逐页目视检查；这不是 PowerPoint 保存关闭重开或专业事实审阅证据。
- `reference-plan.json` 与 `reference-deck.json` 是另一个八页候选：计划引用三份冻结的 NASA HTML 快照，图片引用官方 2K 原图附件，逐页声明来源与待审阅结论。`generate-candidate.mjs` 可重建它们。构建版 Taskpane、Electron PC 与 Rust Relay 的本机故障注入将该图导入 PC 附件并生产八页，回读 PPTX/PDF 通过；这仍不是用户在真实 PowerPoint 中完成任务的验收。

## 受控故障与任务提示

`fault-scenario.json` 定义一次首选 4K URL 超时、再尝试官方 2K URL、核对持久缓存及原任务续跑的记录清单。现有 `apps/office-addin/tests/presentation-p0-14-fallback.test.ts` 使用本包真实 2K PNG 做跨插件/PC 故障注入：测试专用下载适配器把经过公开 URL 检查的占位地址映射到 localhost，以便确定性注入超时；**不把 localhost 视为生产允许的外部地址，也不声称真实 NASA 网络发生过故障**。

> 用两份 NASA 原始资料和已授权的 2024 温度异常图制作八页 16:9 中文专业汇报。先核对来源、图像使用条件和图例；在写页前处理一次受控首选图片超时。若官方 2K 候选成功，则记录候选次序、失败时间、最终图片 SHA256、来源署名与缓存命中，再继续原项目和请求，不重复生成已完成页。没有可验证素材时明确提示待补料，不插入空白图。区分 1951–1980 基线异常与绝对温度、地区天气和未来预测。保存逐页 QA、故障日志、最终 PPTX 与 PowerPoint 保存关闭重开及可编辑证据。

运行 `node generate-reference.mjs` 可重建参考稿，`node verify-materials.mjs` 和 `sha256sum -c SHA256SUMS` 核对固定包。真实任务还需人工核查素材许可例外、科学解释、最终截图与宿主回读；候选材料数不能计入 20 项通过数。
