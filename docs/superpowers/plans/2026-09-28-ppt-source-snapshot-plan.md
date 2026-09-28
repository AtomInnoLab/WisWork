# PPT 来源原文快照绑定实施计划

目标：保留原始来源 URI，同时把证据读取绑定到当前文档的文本附件。非目标：自动抓取网页、核验网页与附件同源、事实判断或真实 PowerPoint 宿主验收。

架构：计划来源新增可选快照 ID，统一解析器兼容旧 `attachment:` URI。PC 和插件依该解析器查找附件；交付报告仍显式标记没有快照的外部来源。原始 URI 保留在页脚与报告中，冻结任务和证据摘要纳入快照字段。

约束：附件必须是当前文档所有、摘要完整、摘录逐字可定位；不读取任意文件路径或网址。旧版数据解析行为保持。修改冻结证据要创建新计划修订与任务。

1. **计划与报告协议**：修改 `packages/pptx-engine/src/presentation-plan.ts`、`presentation-claim-evidence.ts`、`presentation-page-reviews.ts`、`presentation-delivery-report.ts` 及对应测试。先写旧计划/新快照/冲突拒绝/交付问题转变的失败测试，再实现解析器；定向与引擎全套测试通过。提交独立协议改动。
2. **PC 绑定链**：修改 `apps/shell/src/main/presentation-production.ts`、`presentation-source-audit.ts`、`presentation-service.ts` 和集成测试。先验证外部 URI+快照的整稿/逐页编译前检查、原文证据窗口与文档隔离失败，再实现。提交 PC 改动。
3. **Office 插件与完整路径**：修改 `apps/office-addin/src/skills/powerpoint/presentation-project.ts`、`presentation-planning.ts`、必要的 Agent 说明与测试；增设 URL→上传快照→修订计划→新任务→报告的跨层集成测试。验证旧版状态兼容与新字段不被插件误丢；提交插件和集成改动。
4. **完成检查**：运行引擎、Shell、Office 全套测试与类型、lint、构建；核对本阶段文档与实现，更新整体进度台账。由于没有真实 PowerPoint 和真实专业材料，不提高未跨档位的百分比，不将这批标记为全方案完成。
