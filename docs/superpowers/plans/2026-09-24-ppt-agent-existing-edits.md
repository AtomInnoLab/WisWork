# 现稿原生修改、保存点、撤销与局部复核

基线 b7189017，沿用隔离分支。执行原方案§6.1–6.4、O5，实机验收仍暂缓。原方案已获批准，不重复设计审批。

## 目标与约束

将已读取的当前页/选区基线用于现稿单对象文字/几何修改。确认前重读基线，写前持久保存点，写后对象回读；重开后可按记录检查、继续或撤销。截图和视觉复核只针对变更页，记录绑定具体截图；不伪造生成项目或导入回执。复用现有提案、原生按ID写入、统一历史容量和未决保护。不新增依赖，不改变自动确认策略，不重建页面。

边界：本批单对象单操作，非多操作原子ChangeSet；文字修改/撤销只恢复文字内容，不承诺恢复全部富文本run；不对图表/图片/组合进行文字重写。范围仍来自基线，ID不按页索引替代。QA历史结果不等于当前宿主验收，手工修改需新截图复核。回退旧代码对未知历史类型拒绝读取，保留用户记录，不删除来降级。

## 共享记录与接口

新 presentation-existing-change.ts：PresentationExistingChange 为text/geometry判别联合，字段 version:1, changeId, documentId, baselineId, baselineDigest(SHA256), scope:{slideIds:string[],shapeIds?:string[]}, hostSlideId, shapeId, shapeType, kind:'text'|'geometry', before/after（kind=text为string，geometry为四个pt数字）, state:'pending'|'applied'|'undo_pending'|'undone'。可选review:{screenshotDigest,capturedAt,reviewedAt,status:'pass'|'fail',notes}，仅终态可存、状态变化清除；是历史截图审查记录。不可变核心字段严格校验，记录最多192KiB，review最多8KiB。目标必须属于scope，ID与文本/几何沿用已有边界。

统一PresentationHistoryEntry新增kind:'existing'，record为上述记录；新slot wiswork.presentation.existing-change.v1，history.heads.existing，独立历史ID existing:<changeId>。新readExistingChange(changeId:string)精确选ID无fallback；writeExistingChange(record,expected:record|undefined)与history同次settings.save，复用保存失败回滚、身份检查、序列号、64条/1MiB和跨类型未决保护。预算预留终态/撤销状态增长及完整review空间。现稿不含projectId/requestId/artifactDigest，生成稿筛选排除它。

baselineSkill增加snapshot(baselineId):DeckBaseline|undefined（副本）；基线校验通过既有check工具，写入前后用基线ID、文档身份和epoch校验，不因内容已修改而伪造旧基线仍有效。

## A：持久记录与统一历史

拥有presentation-existing-change.ts、presentation-change-history.ts、presentation-document.ts、presentation-history.ts及新的existing-change-binding测试。扩展summary映射但生成成果工具仅筛选生成记录。保存新记录pending，精确CAS；后续状态严格pending→applied→undo_pending→undone，同状态仅review变更，不改核心。新记录不能带review，review只能在applied/undone；重复原样保存幂等。保护未决期间跨类型新事务，重开/SaveAs/损坏/容量/旧版分歧均有回归。TDD并限定提交；其他文件类型影响由Root/B整合。

## B：工作台支持现稿

拥有agent/presentation-changes.ts、presentation-changes-card.tsx与相关tests。Options新增existingAvailable?:()=>boolean，复用listChangeHistory/documentId；离线/无artifact时仍能查看当前文档existing条目，原生成条目继续按原artifact和available筛选。现稿row.kind可映射text/geometry，新增source?:'existing'标识，scope和review明确历史性质。按ID分发 inspect/undo/resume_existing_presentation_change，输入只change_id，不塞project/page伪身份。动作前后重读精确记录与doc/availability，保留clear/取消/忙碌保护；未决记录可检查/继续，applied可检查/撤销，undone可检查。TDD限定提交，不改runtime。

## Root：提案、恢复、截图复核与接线

新增presentation-existing-editing.ts + tests，修改baseline、host-runtime与跨层tests。工具edit_existing_presentation_text/geometry以baseline_id/slide_id/shape_id和after值发起；list/inspect/undo/resume_existing_presentation_change按记录操作。复用原生expected-before写入和proposal validate/execute/verify，写前pending持久化；恢复先比较当前值是before还是after，已完成宿主写只补回执，未知值拒绝。撤销必须先保存undo_pending，再原生写回before；同样支持回执失败继续，不依赖会话基线。

capture_existing_presentation_change获取当前目标页截图，校验记录状态/目标值/文档；record_existing_presentation_change_review必须关联本会话capture及重新截图匹配，才持久保存历史视觉判断。记录未来变更时不声称旧review仍代表当前页。工具输出原文为不可信内容。写入通过既有QA失效hook仅标记该宿主页，工作台刷新与清会话均接线。

真实binding + runtime + proposal +模拟host：现稿基线→文字/几何确认→重开→历史→撤销、回执失败恢复；确认间手工修改/另存/改选区拒绝，未决阻止跨来源新事务，局部截图review不冒称全稿QA通过。

## 验证与交付

独立全diff审查并修复重要问题；全仓test/typecheck/lint/licenses/format/diff、Office Add-in和Shell构建。真实验收仍0/20，不因工程测试虚增。阶段报告更新整体百分比、已完成和下一步。限定本地提交，保留工作树，不合并/推送/部署。
