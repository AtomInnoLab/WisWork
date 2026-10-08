# 复制页面持久事务实施计划

基线 ffdfc8cf，按已批准原方案 §6.4、§12.3、§14.2，继续同一隔离实现分支和本机 PC 原页备份授权。

## 目标与架构

保留 duplicate_slide 与 execute_office_js 中单项 duplicate_slide 两个现有入口，接入写前保存点、持久未决与可撤销记录。复用 BrowserPresentationPageReplacementAdapter.stage/discard 的原生单页 KeepSourceFormatting 插入、严格页序/内容核对和写前守卫；不会调用移除源页的 commit。existing_batch 增加独立第四版本，明确记录源页、新页身份及操作回执，旧版本保持读取兼容。

所有操作使用明确的提案确认；不确定写入不重放。只读 inspect 给实际包/页序证据，明确 reconcile 提案只认领唯一相邻新页且原包及新包一致，不能由无变化推断写入成功；undo 仅移除可证明属于此变更且未被修改的新页，源页和最终原页序必须仍相符。源页变更或复制页变更时不删除用户内容，保留待核对记录。实际原页序/原包证明完全未创建新页时，可经明确确认将未决记录闭合为未应用的 undone，不重放写入。源文档上限512页，插入后的513页使用现有临时导出和移除适配器核对；不宣称其它读取入口已支持更大文档。

## 交付单元

1. **复制账本与技能**：presentation-existing-batch.ts 增加严格版本/转移/预算验证；新增 presentation-slide-duplication.ts 与专项测试。接口 propose(slideIndex,explanation,signal,toolName)；工具 inspect/reconcile/undo_slide_duplication、capture/record_slide_duplication_review。依赖实际 documentId、available、PowerPointAdapter.exportPresentationPagePackage、既有 page replacement adapter、PC request、proposals 及 read/writeExistingBatch。备份读验与账本未决先于写入；每次 await 后核文档/能力/epoch/CAS，实际适配器最后写前核。RED 覆盖备份失败、断连、ACK不确定、参数引用、重开核对、明确认领、内容/页序漂移和撤销后保存失败，GREEN 后保留可复查证据。
2. **入口与工作台**：root 修改 powerpoint-skill.ts、host-runtime.ts、history/existing-editing 投影、changes/card 及 ACP。替换两个旧内存复制分支，能力不足拒绝；工作台给 inspect、明确认领和 undo 入口，正确区分源页/新页与历史截图。实际 Runtime 测试覆盖确认前不写、原页 PC 保存、持久写前意图、重开和确认撤销。旧夹具迁移保持原断言。
3. **独立审查和交付**：审查完整单元及变更范围，处理严重/重要问题，完整八工作区回归、Office/Harness 类型、定向静态与生产构建。源码提交、构建 ID、阶段报告与严格百分比台账关联。

## 全局约束、回退与发布

不新增依赖、迁移、部署或清理真实用户数据。原方案不改；合成SDK验证不作实机验收。写后截图可明确采集并复核，实际包匹配不能认证视觉质量。复制只失效已证明源页的QA，撤销覆盖源页与已知复制页，明确元数据认领不失效QA；未知旧入口保持保守范围。回退不能忽略已存在的第四版本账本，需保持兼容读取或先恢复；不能以回退自动移除页面。灰度/专业/宿主退出继续按原条件核证。
