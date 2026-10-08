# 统一变更历史与逐项撤销

基线ebcd9aa3。沿用隔离分支codex/ppt-agent-implementation，按原方案§6.2/6.4、O5继续实现；实机验收0/20暂缓。用户已批准原方案及持续实现。高保障：持久化兼容/多key同次保存/撤销，TDD、分工、独立审查、全仓验证。

## 目标和方案

现有文字/几何单slot会覆盖旧记录。新增统一有序历史并与原保存点在同一次Office设置save中保存，原生写入仍在保存成功后执行；失败一起回滚，不允许单边成功被当作已完成。文字/几何支持按change_id读取和撤销较早保存点，图片使用原有记录键，整页旧终态记录仅历史展示。所有动作继续复用原工具的当前值/身份/确认/QA核验，不直接把旧快照写回宿主。

新增统一ChangeSet摘要为单操作描述：scope、intent（操作描述，不伪称保存了用户原话）、operations、preserved、validation（要求，不是通过证明）、risk。不实现多操作原子批次、自动低风险放行、redo、自动跨请求/对象ID重映射或任意现稿DeckBaseline。历史逐条撤销不是可越过依赖的全局Undo栈：对象ID/内容/顺序已变化时保守拒绝。

## 共享接口

新文件presentation-change-history.ts：

- PresentationHistoryEntry判别联合：{id:string, sequence:number, legacy:boolean, kind:'text'|'geometry'|'image'|'page', record:对应记录}。
- 历史最多64条，总序列化字节+未决/后续撤销预算<=1MiB。id由kind与记录身份确定，文字/几何/整页changeId；图片使用完整任务/页/原图身份（不把可碰撞的显示名称当ID）。sequence首次加入时分配严格单调整数，状态更新不改变sequence。
- validatePresentationHistoryEntry(value):boolean/typeguard；导出historyEntryId(kind,record)供controller校验。legacy=true标注首次导入旧slot、顺序未知；不伪造历史发生时间。旧记录不得静默丢弃。
- binding.listChangeHistory():PresentationHistoryEntry[]（副本）。readTextChange(changeId?:string)、readGeometryChange(changeId?:string)，省略保留原最近slot语义，显式ID按历史精确查找，不存在返回undefined，绝不fallback到最新。原write接口不变，较早ID状态更新仍以expected记录做CAS。

## A：持久历史与记录选择

拥有presentation-change-history.ts、presentation-document.ts及新增/相关binding测试。现有四类write接入同一历史key wiswork.presentation.change-history.v1。历史envelope保存entries及text/geometry/page当前head引用；images与原map校验一致。读历史核对原slot/head，旧版本写了slot而没有同步历史时拒绝旧历史操作，不猜测同步。头指针可以指向最近操作的较早记录；不以sequence最高替代精确head。旧key不存在历史时只读映射旧记录；首次写原子导入旧记录并加入/更新新记录，legacy顺序明确未知。所有head必须引用同kind条目，唯一identity/sequence严格验证。

writeText/Geometry更新时：如果snapshot.changeId与expected相同，从历史精确读取CAS；新ID仍与现有slot做CAS且须pending起步。保留不可变内容/合法方向/SaveAs/损坏封锁。历史保存不允许新事务绕过未决记录；同一记录可继续原方向恢复。仅历史容量达到上限时明确失败，不驱逐未决或已完成记录。图片预留复用现有函数；text/geometry状态字段留裕量；整页保守预留其192KiB记录上界，保证不在插入/删除后才因统一历史满失败。原四类各自预算继续适用。避免历史校验与原read递归，使用私有raw readers。

同一次settings.save完成legacy对应key与historykey，核验文档身份和两个key；失败回滚两个，无法证明回滚则fail closed。不另起后台异步镜像保存。原其他receipt/QA逻辑不重构。

TDD：多次同类不覆盖、跨类型顺序、精确ID/CAS、重开、迁入旧数据未知顺序、同存储失败回滚/损坏/SaveAs/旧版slot变化、容量写前阻断、pending不被新变更绕过。限定提交。

## B：文字/几何按记录撤销工具

拥有presentation-page-editing.ts和对应page-editing/text-change测试（不改binding）。更新options.readTextChange/readGeometryChange可选changeId参数；read/inspect/undo/resume两类工具schema及输入接收可选change_id（^[A-Za-z0-9_-]{1,128}$）；只有这些历史工具允许该字段，新edit禁止。读到记录后显式检查changeId一致，旧callback忽略参数也不能错误操作最新记录。不传ID保持旧调用契约。后续所有重读/校验固定同ID，写接口expected为该记录。保留page/document/request/source/artifact/host绑定，冲突时拒绝，不隐式选择另条记录。补充systemPrompt说明列出记录再按ID撤销、每次仍确认、历史不证明当前状态。

TDD：两次修改后撤销后一次，再撤销前一次；跨页/任务/不存在ID/忽略参数的旧reader拒绝；较早pending恢复不触及较新slot；输入边界。限定提交。

## Root：工作台、统一摘要工具与真实跨层

host-runtime.ts传递可选参数/listChangeHistory，runtime.changes优先历史（无接口fallback旧聚合）。presentation-changes.ts读取全部记录并按现有artifact/source/digest/page过滤，顺序按sequence降序；legacy显示顺序未知。较早text/geometry action附change_id，兼容旧fallback只按page；page已过期终态无动作，最新page仍走现有工具。动作前重新核对选中原始记录与上下文，旧响应/clear/busy保护不变。

新增presentation-history.ts skill注册list_presentation_changes，输入project_id?，只输出当前任务限量摘要，包括change_id（适用时）、kind/state/order、单操作ChangeSet scope/intent/operations/preserved/validation/risk；不输出全部图片字节或长全文。UI卡片展示作用域/操作摘要/风险/验收要求与历史容量边界；仍不自动确认或宣传已QA通过。

跨层：真实binding+runtime+proposals+模拟宿主，文字1→几何→文字2→重开→按记录撤销文字2→几何→文字1，验证各保存点仍在、目标QA局部失效、冲突拒绝、完成回执失败按正确记录恢复且不重复宿主写；当前请求工具列表不泄漏其他任务。旧无history接口兼容回归。

## 验证与兼容

独立审查全diff并修复重要问题；全仓test/typecheck/lint/licenses/format/diff，两端构建。原方案不变，报告准确区分单操作ChangeSet摘要与完整多操作ChangeSet能力；记录总进度百分比和真实验收0/20。保留备份/历史数据，无自动删除、外部依赖或后台迁移。降级版无法维护新历史，升级后发现分歧封锁历史写，不声称降级后仍支持新历史；不删除用户数据来解锁。仅隔离分支提交，不合并/推送/部署。
