# PPT Agent 修改后复检

基线935e1920。沿用已批准方案的Office.js原位修改与分层QA设计，现有隔离分支继续，不合并或部署。

## 目标与边界

复用文本/图表/OOXML等现有修改提案，在确认执行前持久化需要重新检查的状态。修改与QA采集互斥；失败、取消或未知写入也不恢复旧的通过结论。再次capture与review沿用已实现的精确页映射和截图一致性检查。

本轮不新增编辑器、自动修版、撤销体系或人工宿主验收；不监听用户手工修改事件。对插件发起的PowerPoint修改保守标记文档内全部已有QA页面需复检，避免对不透明OOXML/script范围猜测。

## 架构与契约

- StructuredProposalController可选第二参数hooks：beforeWrite(proposal, signal)返回Promise<void>；afterWrite()同步释放。validate成功后、execute前进入；beforeWrite返回后重新validate，防止保存期间的用户修改被覆盖；一旦进入beforeWrite，无论其失败或execute/verify失败，finally都释放。拒绝/validate失败不触发。
- PresentationQaRecord页增加可选recheckRequired:true。保存旧结构与Agent判断作为历史；capture创建新页记录时清除标记。UI优先展示需重新采集，不将旧pass展示为当前通过。
- QA skill增加beginMutation()/endMutation()。忙时拒绝开始修改；开始修改时清除live授权并递增epoch，修改期间拒绝QA工具；clear不能提前释放正在执行的修改锁。
- DocumentBinding.invalidateQa()与导入/QA保存共用设置队列，原子标记所有已有页；无记录无需保存；保存失败回滚本地旧值并阻止文档写入。
- runtime将仅PowerPoint提案接入hook，文档状态标记保存成功才调用原有修改；结束后刷新QA卡片。现有宿主工具保持原有预览、确认、校验与回读。

## 可独立验收单元

1. 提案生命周期：proposal-controller.ts及测试，涵盖确认时顺序、拒绝/校验失败不触发、beforeWrite失败禁止写入、取消和execute/verify失败最终释放。先红后绿，独立提交。
2. QA模型/锁/UI：presentation-qa.ts、presentation-qa-card.tsx及测试。旧record兼容；标記strict校验；历史结论保留且显著标记；写入期间禁QA，freshcapture清标记；先红后绿，独立提交。
3. root设置/运行时整合：presentation-document.ts、host-runtime.ts、跨层测试。全量历史标记持久化/回滚、实际提案修改前失效、失败不复用旧live、成功修改后重新capture/review可恢复。独立提交连同报告。

## 风险与回滚

保守失效可能要求未受影响页也重新检查；先保证不误用旧结论，后续有可靠变更范围时再按页缩小。持久化新字段为可选；旧版严格校验会拒绝带标记的新记录，所以不自动降级移除标记以免恢复虚假的当前通过。代码回退时应保留文档历史，不保证旧客户端能读取新标记，不作自动数据清理。

## 验证与收尾

独立审查全部diff；定向并发/失败回归；全仓test/typecheck、变更lint、插件与Shell构建。报告明确mock与Office实机边界；保持工作分支，不部署、不推送。

## 审查修正

- 保存复检标记增加异步窗口，因此在hook后再次validate并检查取消，文稿发生变化时阻止执行。
- 原业务记录64KiB/聚合256KiB预算保持；每个合法recheckRequired:true的固定JSON字段字节单独允许。先用8*32个标记的最大空间限制原始数据，严格解析后只扣除实际合法标记，兼容已存近上限历史并保持硬上限。
