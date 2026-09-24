# PPT Agent：修改差异与撤销工作台

基线c349abb9，既有隔离分支。按原方案§6.2/6.4、O5推进，实机0/20继续暂缓。本轮使用已批准方案，跨持久化/原生修改/工作台接口按高保证实施：TDD、独立实施单元、全diff审查、全仓验证。不修改原方案。

## 目标与边界

补齐稳定业务页文字修改的持久保存点、单级撤销和中断恢复；把文字、几何、图片、整页替换的现有记录汇总成实际可操作工作台。记录不是完整ChangeSet历史：文字/几何各最近一条，整页替换最近一条，图片沿用现有32条上限。图片暂无撤销，不伪造按钮；整页和图片仅显示身份/摘要差异，不冒称视觉diff。任意现稿基线、统一跨类型多级撤销继续未完成。

复用现有document绑定/串行settings队列和proposal确认/QA失效。所有动作工具检查当前文档、artifact/request/source、稳定页面映射、对象当前值；不由UI直接写Office。未知中断状态只能检查，不自动重放。UI读取历史不读取或写入宿主，操作按钮才调用对应工具；会话/文档/任务切换取消旧操作结果。

## 接口与交付单元

### A文字保存点与恢复（agent）

新增presentation-text-change.ts导出PresentationTextChange/validatePresentationTextChange；字段与PresentationGeometryChange一致（version,changeId,documentId,projectId,requestId,source?,artifactDigest,pageId,hostSlideId,shapeId,state），但before/after为string各<=12000，完整记录<=192KiB，严格字段及ID/摘要验证。状态pending→applied→undo_pending→undone。单slot，只有旧终态可新建，不覆盖未知未决；同changeId除state外不可变。

presentation-document.ts新增readTextChange()/writeTextChange(record,expected) CAS、快照、队列、SaveAs/保存失败回滚，私有key wiswork.presentation.text-change.v1；不迁移既有记录。另暴露listImageReplacements():ImageReplacementRecord[]，由现有readImageRecords校验后返回副本，供工作台只读枚举；不新增图片存储。

presentation-page-editing.ts新增可选readTextChange/writeTextChange配置，edit_presentation_page_text在可用时写前pending、回读后applied，不可用保持旧路径。新增read/inspect/undo/resume_presentation_text_change，输入同geometry工具{project_id?,page_id,explanation?仅resume}。恢复分类ready_to_apply/already_applied/manual_review/not_pending，当前值必须完全等于before或after；未变更文字不建立模糊保存点。undo必须当前等于after，先保存undo_pending再写before；resume保存方向不可改，已在target只补回执；所有写需要proposal确认、二次观察和adapter expectedText。取消写后仍核对实际结果，保存失败留下可检查的pending。不因截图通过掩盖中断状态。

A拥有上述record、document、page-editing与相关tests，限定提交。可提取最小复用逻辑但不重构无关图片/几何。先RED：持久重启/CAS/损坏/SaveAs/限额；文字正向/撤销/两方向恢复/手工冲突/保存失败/取消、旧无storage兼容。

### B工作台controller与view（agent）

新增src/agent/presentation-changes.ts和presentation-changes-card.tsx及tests，不编辑App或host-runtime（root负责）。

createPresentationChangesController(options): {snapshot(),subscribe(),refresh(),run(entryId,action),clear()}。
options={available():boolean,artifact():CompiledPresentationArtifact|undefined,documentId():Promise<string>,readTextChange?(),readGeometryChange?(),readPageReplacement?(),listImageReplacements?(),executeTool:AgentSkill['executeTool']}。
snapshot={phase:'idle'|'loading'|'acting',projectId?,requestId?,entries:PresentationChangeEntry[],notice?,error?}；entry显示id、kind(text/geometry/image/page)、pageId/state、before/after字符串、actions。具体类型在开始即发给root。每条内部保存完整record指纹，run重新读取和比对当前原始记录；stale不执行。refresh严格validator+doc/project/request/source，text/geometry须artifactDigest一致，page允许当前请求等于parentRequestId或requestId，且仅productionartifact；image同doc/project/request/source。只显示当前artifact对应记录，不自动切换任务。持久坏记录明确错误不假空。

动作：text/geometry applied→undo，pending/undo_pending→inspect/resume，undone无操作；image pending→inspect，只有baseline+newShapeId可resume，complete不可undo；page pending→inspect，inserted→inspect/resume，staged→inspect/commit/discard，discard_pending→inspect/discard，commit_pending→inspect/commit，applied→inspect/undo，undo_pending/restore_inserted→inspect/undo，discarded/undone无操作。内部映射对应既有工具；text/geometry{project_id,page_id}，image加shape_id:oldShapeId；page{project_id,change_id}。工具自身再次验证/生成proposal；不直接confirm。返回结果只显示固定安全提示或经过验证的状态码，不显示不受控tooloutputHTML。变更后需要重采集QA提示明确。

view props {controller,disabled?:boolean}，真实点击refresh/inspect/undo/resume/commit/discard。busy/disabled禁用动作，差异保留安全文本、长内容可折叠且显式标明。显示scope与“最近保存点/不完整历史”限制，图片/整页差异类型明确。支持未恢复任务/空记录/error。先RED组件真实click、controller旧响应/clear/docchange/taskchange/指纹变化/非法action/隔离/无宿主写/工具失败。限定提交。

### Root接线与跨层

host-runtime类型和实例接入read/writeTextChange、listImageReplacements；工具路由注册4个text工具、QA稳定页白名单加入undo/resume text。实例化changes controller，执行入口只分派pageEditing/pageReplacement；每个相关工具返回、持久record写后及artifact切换后刷新，clearSession清空并取消。不要每次refresh反向触发notify循环。全局proposal执行后记录通知可通过write包装刷新，终态失败也发布状态。App ui.changes=runtime.changes，添加card并沿用busy/applying/proposal/upload/projectPhase禁用条件。跨层真实binding+runtime+proposal+模拟原生文本：修改→保存→重开→工作台撤销→确认→QA局部失效，以及写后存储失败→检查→只补回执，手工改动阻止撤销。无实机声明。

## 验证、兼容与回滚

先定向RED/GREEN；完整diff独立审查，最多两轮修复；全仓test/typecheck/lint/licenses/format/diff，测试结束后两端build。无外部依赖，文本新增独立key不更改旧记录；降级版本不会显示文字保存点且无法提供其持久撤销，不能宣称降级后仍受新pending保护。保留数据回滚、不合并/推送/部署。原方案进度47%基线，O5依据完整主要链路而非单功能数量复评，实机0/20。
