# PPT Agent：逐页导入检查点与安全继续

基线 f1c321a；沿用已批准方案 O3 的写入事务/恢复要求。本轮实现 Office 导入阶段逐页记录；不宣称 PC 编译已变为按页独立生产。

## 路径与边界

PptxGenJS 完成后从实际 PPTX presentation.xml 提取源页 ID → PC 回执保存可选 pages 元数据 → 插件恢复生成物 → 用户确认一次 → 每页保存 inFlight → Office 单页插入并回读 → 保存 completed → 下一页。

依据 Microsoft 官方 sourceSlideIds API：
https://learn.microsoft.com/en-ca/office/dev/add-ins/powerpoint/insert-slides-into-presentation

使用规范 sourceSlideId `<numeric-id>#`，不假定生成器的内部编号。metadata 缺失的旧生成物沿用既有整稿导入。

## 持久状态

现有导入记录增加 checkpoint：version、artifactDigest、sourceSlideIds、baselineSlideIds、completed(sourceSlideId/slideId)、可选 inFlight(sourceSlideId)。完整 base64 的 SHA256 绑定来源。旧 pending 仍保守阻止重复，旧 complete 仍防重复。

- 写入前 checkpoint 落盘失败：不调用 Office。
- 已完成页不重复；中断且没有 inFlight 时，实际页 ID 序列必须与 baseline+completed 完全一致，重新确认后只继续剩余页。
- 任一页写入结果不确定或回读失败：保留 inFlight 并停止，不猜测成功、不盲目重插、不自动删页。
- 每页完成记录落盘失败：恢复之前的保守记录，或将本地绑定标为不可信，不能把未持久化的成功作为恢复依据。
- 取消与写入并发时仍回读该页；若确认写入成功，先保存检查点，再停止后续页。

## 实施单元

1. 编译器/PC：真实源页 ID 提取、pages 元数据和持久恢复，保持旧回执兼容。
2. 页面执行/记录：严格 checkpoint 校验、逐页确认执行、部分完成继续、未知写入阻止；状态读取工具与纯摘要。
3. 插件集成：Office insertPage、生成元数据验证、运行时工具/状态、Taskpane 导入记录显示。
4. 跨层测试、独立审查、全仓测试与类型检查、lint、两端构建；阶段小结。

## 验收

覆盖逐页成功、完成两页后取消并继续、源包/文档/页序改变拒绝、未知写入不重复、设置保存失败、旧整稿导入路径、真实生成源 ID 与 Office mock 选择器对应。真实 Office 宿主的保存重开与视觉验证仍需实机完成。

## 执行结果

- PC/source metadata：ee70a05；页面检查点与设置持久化：f15717c。
- 跨层实测：真实8页生成→导入2页→安全中断→重建PC/插件→导入剩余6页，无重复。
- Office单页选择器与回读使用mock验证，尚非真实宿主测试。
- 独立审查无重要运行时发现；测试代码的取消API和严格类型问题已修复。
- 全仓5810项Vitest测试、全仓类型检查、变更lint/diffcheck、插件和Shell生产构建全部通过。
- 保留隔离分支；详细边界与下一步见 product 目录逐页导入恢复小结。
