# 主张与来源时点元数据预检

基线：1952dd40。依据原方案 §4.3 与专业来源策略；整体成熟度42%，实机0/20按用户要求暂缓。沿用已批准方案及隔离实现分支。

## 目标、架构及边界

在现有冻结页面内容预检增加时点元数据待处理项，复用 production_content_check / check_presentation_page_content，不新增工具、存储或依赖。主张 asOf 去除首尾空白后非空时，逐一比较已绑定来源：来源 asOf 缺失或仅空白报告 source_as_of_missing；两端仅规范化空白后仍不同报告 source_as_of_differs。主张未指定时点时不猜测任务要求，不从自然语言 brief 提取期限。保留原始 asOf 不修改计划。

所有差异只表示元数据标记不同，需要核对，可能是合理的多期比较；不把字符串不同认定为过期、冲突、支持不足或事实错误。相同标签也不代表日期有效、语义相同或时效通过。保持 checks.timeliness = not_verified 及其余保守标记，不执行公式、不联网、不改变宿主或复核记录。

## 实现单元

1. 共享契约与算法：packages/pptx-engine/src/presentation-content-check.ts 及 tests/presentation-content-check.test.ts。两个新 code 均必需 claimId/sourceId；兼容既有无新 finding 报告。每页32主张、3来源，最多352项 findings（原最多256 + 96时点项），严格上限，保持未知字段/代码/重复/归属拒绝。TDD覆盖缺失、空白、不同、相同、不指定主张时点、其它页面隔离、不修改输入、最大合法报告及非法归属。限定提交。
2. 插件和跨层：apps/office-addin/src/skills/powerpoint/presentation-production.ts 添加准确建议文字并将 contentCheck 的响应上限设为256KiB（原64KiB无法容纳最大合法来源归属列表，复用 evidence 上限）。tests/presentation-production.test.ts 覆盖两类建议、无状态改变、合法超过64KiB报告和超过256KiB拒绝；apps/shell/tests/presentation-content-check.test.ts 覆盖冻结计划更新及重启时点报告保持原记录、无写入/编译。
3. 独立审查全改动；全仓test/typecheck/lint、format/diff、licenses、Addin/Shell build。记录真实结果、42%口径与下一步到阶段报告，并追加总体/审计/验收记录。限定提交。

## 回滚与兼容

无存储迁移，回退不会改变历史。旧插件对新发现代码会保守拒绝报告，应配套升级；不伪装通过。缺少过期阈值和完整专业时效政策，仍不声明时效验收完成。实机、合并、推送、部署继续暂缓。

## 完成记录

共享引擎、插件建议/传输边界、冻结计划跨层回归已完成。独立审查、全仓测试/类型/lint/格式/许可证和两端构建通过。详见 `docs/product/wiswork-ppt-agent-as-of-precheck-progress-2026-09-24.md`。保持整体42%、实机0/20暂缓，不合并或推送。
