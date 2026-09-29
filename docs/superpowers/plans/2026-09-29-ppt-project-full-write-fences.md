# 演示项目全写入口栅栏：下一工程批次

## 目标与当前状态

依据原方案 §13.1 和 [保留期存储核对](./2026-09-29-ppt-project-retention-storage-inventory.md)，接通尚未受持久生命周期版本约束的项目入口。本文是可实施计划；代码保持当前冻结状态，不表示全入口栅栏、统一删除、定时保留期或用户确认界面已经完成，不提高整体 64% 完成度。

现有可复用基础：`PresentationLifecycleStore` 的独立控制命名空间、固定 revision 的 write/read lease、`PresentationStore.projectScope()` 的 metadata-only 归属证明、production 每次同步写 guard、后台 worker 的 abort/drain。`production_*` 已接线；纯读取缺少控制记录时保留 absence，不隐式初始化，后续出现任何控制记录会使旧读请求失效。`production_job_status` 可能追加恢复事件，仍按写入口处理。

架构：继续复用现有进程内项目锁和生命周期 CAS，不引入第二个调度系统。入口在首个 await 前固定请求、实际 scope、生命周期 revision 或 absence；子模块围住最终真实写入，返回私密结果前复核原版本。删除冻结不等待长项目锁，取消/等待在锁外完成，实际清理须在全部相关工作结束后取得项目锁。

### 全局约束

- 只操作本地临时目录与合成数据；本批不执行真实项目删除、网络部署或真实用户目录清理。
- 同一 projectId 的 deleted 控制记录不可因正文目录消失、换 documentId、首次 save_plan/compile 或研究重试而复活；新项目需新 ID。审计保留期的未来清理不得删除这项最小 tombstone 保护。
- 不续租：旧请求不能在 await 后重新 capture 最新 revision。等待项目锁前就固定原 revision，取得锁后仍检查它。
- write lease 与 read lease 分离。纯读没有 assertWritable；缺控制记录不写文件、不创建目录、不修复账本。只读路径如误触写方法，显式 access_denied。
- 保留现有输入/输出/ZIP/资产预算与幂等回执，不降级测试断言、不靠内存 fallback、不扩大原权限。PC 单主进程检查不宣称跨 OS 进程原子 CAS。
- 用固定有限错误码；异常 message 只读取一次，原始路径、正文、URL、令牌不得穿出错误边界。
- 全局品牌、文档附件/acquisition history、existing/master/package backups 没有项目独占证明，不按“请求所在项目”阻断整个文档、绑定成独占资源或删除。这里只冻结明确的项目写入/项目引用。

## 实读入口与归属

路径均相对仓库根目录；下列函数名为实施定位依据，行号不作为合同。下文 Shell 源文件 basename 均指 `apps/shell/src/main/<basename>`，Shell 测试均指 `apps/shell/tests/<basename>`。

| 入口 / 当前函数                                                                          | 实际 scope / 是否允许首次创建                                                                       | 真实写入或只读性质                                                                                                                                                |
| ---------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `presentation-service.ts`: save_plan / `store.savePlan`                                  | 已验证 plan.projectId + documentId；首次 expectedRevision=0 可建立新项目                            | research 校验 await 后写 plan 与修订；必须先创建控制记录，不能先写正文再补 lifecycle                                                                              |
| 同文件 compile / `store.begin`, `store.complete`；resume                                 | deck.id/projectId + documentId；合法首次 compile 可无 plan，resume 只能已有 receipt                 | 首次 begin 及编译 await 后 complete 都写；已 compiled 的重试返回仍需原 read/write admission 检查                                                                  |
| accept_plan / set_plan_page_lock                                                         | 只允许实际已有项目/计划                                                                             | 同步 `acceptPlan` / `setPlanPageLock`，目前主项目锁有但无生命周期 fence                                                                                           |
| audit_sources / `handlePresentationSourceAudit`                                          | 实际已有计划项目                                                                                    | beginSourceAudit、成功 finish 和 catch 中 failed finish 都写；read_source_audit 才纯读                                                                            |
| get_plan / get / status / read_import_source / export_pdf                                | 项目读取；不为读取创建 lifecycle                                                                    | status 会 await research summary 与附件审查；PDF 绕过主锁并 await render/load/copy/save；返回 body/字节前须原 read lease 复核                                     |
| comment_add/resolve 与 `PresentationCommentLibrary`                                      | 入口已经读实际 store.plan 证明 scope；不创建项目                                                    | add/resolve 同步写；comment_list 纯读且不得初始化控制记录                                                                                                         |
| manual_observation_begin/complete/delete；preference_save_observation                    | observation 的已验证 project/document；原 API 支持无正式 plan 的记录                                | begin/complete/delete 写独立库；saveObservation 以实际 observation.projectId 为目标，不能仅信顶层 request                                                         |
| preference_save/import/delete 与 `PresentationPreferenceLibrary`                         | save 的 projectId 在已解析 preference 内；import 的目标 projectId 与 source scope 不同              | save/import/delete 写；get/list 纯读。导入先以源 read lease 验源，再以目标 write lease 写目标；不能把 source 当目标归属                                           |
| `createPresentationResearchService`                                                      | document+project，研究可先于正式 plan 建立                                                          | build/abandon/delete 写；capabilities 无项目数据；list/latest/read/delete_status 为只读。不得因 PresentationStore.projectScope 缺失永久禁掉合法 pre-plan research |
| `createPresentationTeamService`                                                          | teamId 先经 ACL 授权读取 ledger，再取其真实 documentId/projectId；create/plan_read 用已验证实际计划 | create、publish、member set/revoke、comment add/resolve 写；identity 不属项目；project_read/plan_read 纯读。context/expectedIdentity 不是磁盘归属证明             |
| `createPresentationDeliveryBundleService`                                                | 实际 production record + document/project/request；不创建项目                                       | begin/chunk/finish/delete 写；目前 list/metadata/read 也会 mkdir/cleanupTemporary，metadata/read 还可能 save(receivedBytes)，不能直接称纯读                       |
| `createPresentationPageBackupService`                                                    | project + document + 实际 child/parent production lineage                                           | begin/chunk/finish 写；status/read 当前也先 directory(create)，需消除纯读 mkdir。不能丢 parent/input digest 检查或原 lost-ACK overlap 语义                        |
| existing_page_backup_* / master_backup_* / package_backup_* / attachment_* / brand_kit_* | 文档、change 或全局共享 scope；不是 project scope                                                   | 本批不凭请求上下文补猜测 project ownership。待显式引用/事务归属单元提供证据，保持原协议与恢复可用性                                                               |

## 单元 1：新项目 admission 与前台工作登记

**所有权**：1A 为 `apps/shell/src/main/presentation-project-write-lease.ts` 与新增 `presentation-project-creation-lease.test.ts`；1B 为 `presentation-jobs.ts` 内的前台登记与新增 `presentation-project-work-stop.test.ts`。两份独立可验收交付，不编辑 service.ts，输出接口供主集成者使用。

1. 保留两个已有 capture 函数语义。新增明确的 `capturePresentationProjectCreationLease(...)`，返回现有 write lease 形状；仅经过验证的首次 save_plan、首次 compile、pre-plan research_build 等创建入口调用。不是用户可传的 allowCreate 标志。
2. 新控制记录先于业务正文，policy 仍 null/null；已有控制记录仅 assertActive，绝不重建或清 tombstone。调用 `initialize` 的既有 project.json 外部归属/目录安全校验必须保留；已有项目优先 `projectScope()`，不扫描生产历史正文。
3. 创建入口先完成预算、ID、嵌套模型与创建条件校验。无项目且 save_plan.expectedRevision≠0 不先造控制记录；已有版本按原业务 CAS 拒绝。控制成功而业务失败只留下明确 active 控制，不伪造成功计划。
4. 新增可按 `{root,projectId,documentId}` 停止/等待的前台登记。每 invocation 保留自有 controller、requestId、settled、固定 scope；不能借外部客户端 controller 充作删除执行器可取消的句柄。复用当前进程级登记，允许多个排队 foreground，不用一个 requestId 覆盖另一个。现有 worker/pause 语义不变。
5. `stopPresentationProjectWork(scope)` 在锁外 abort foreground 与现有 workers，并等待实际 settled；普通客户端信号与自有信号合并。编译器没有 signal 接口，不能宣布已中断：等它真实返回，再由 guard 禁止 complete。

**RED→GREEN**：真实 Store 首次控制创建；正文不存在但 deleted 控制已留存仍拒绝；同 ID 换文档拒绝；两个实例同项目排队捕获旧 revision；多个 foreground + background 只 drain 精确 scope；无法中断 compiler 延迟返回、无晚 complete、无锁内等待死锁。1A/1B 各自独立审查通过后做 scoped checkpoint。

## 单元 2：主 Service 计划/编译/审查/只读输出

**所有权**：主集成者独占 `presentation-service.ts`；本单元另拥有 `presentation-source-audit-history.ts` 和新 `presentation-service-full-write-fences.test.ts`；保留现有 `presentation-source-audit-history.test.ts`、`presentation-pdf-export.test.ts`。消费单元 1 接口，其余子服务的 factory 接线也由同一 service owner 后续串行合并，避免多人改主路由。

- 将已受预算约束的非 production 项目请求也在首 await 前深拷贝；保留 hostile toJSON/exception 测试。固定 nested deck/plan、projectId/documentId/requestId 与锁 key，不能只固定几个字符串而继续使用原嵌套数组。
- save_plan/compile 新项目使用 creation admission；已有项目及 resume/accept_plan/set_plan_page_lock/audit_sources 使用 existing write lease。固定版本后才排队主项目锁，取得锁立即复核。
- `save_plan`：research await 返回后及 savePlan 紧前复核；compile：research/附件/编译各 await 后复核，begin 与 complete 紧前复核。失败只保留最后已知 receipt，不在 deleting/取消后补普通 failure 内容。
- `handlePresentationSourceAudit` 加可选同步 assertWritable；begin、success finish、catch failed finish 均检查。错误或取消后 guard 失败直接保留原 running 回执，不把 freeze 映射成可补写 source_unavailable。
- get/get_plan/read_source_audit/read_import_source/status/export_pdf 使用 read lease；absence 保持 absence。status 的各层 catch 不得吞掉 lifecycle 冲突后继续返回正文；PDF 每次 render/load/copy/save await 后和最终字节返回前检查原版本。
- 保留未知项目的原空状态/不存在语义，不为纯读制造项目。legacy 独立库可先于 PresentationStore 存在，正文只能在其自身真实 scope 校验后返回；空结果不作为写授权。

**RED→GREEN**：首次 save_plan/无 plan compile；compile success/failure 在 freeze 后返回均无 complete；排队时修改嵌套 deck/plan、身份及 requestId；取消后 receipt 字节不变；重开 deleted 不复活；audit catch 不晚 finish；PDF/status/source read await 中 freeze 不返回旧正文；纯读前后目录/文件字节相等。保留原完整页序、资产预算与幂等重试断言。source owner 做独立 scoped checkpoint。

## 单元 3：同步评论、观察与偏好库

**所有权**：`presentation-comments.ts`、`presentation-manual-observations.ts`、`presentation-preferences.ts` 与现有 `presentation-comments.test.ts`、`presentation-manual-observations.test.ts`、`presentation-preferences.test.ts`、`presentation-preference-import.test.ts`，新增 `presentation-synchronous-write-fences.test.ts`；service owner 接 factory/callback，不并行编辑 service。

- 保留每个库现有 schema 和 scope 解析。注入同步 write guard，放在创建目录、删除目标以及最终 publish 紧前；库直接调用也能被测试，不能只守 service 外层。
- comment 仍须已有 plan。观察/偏好历史无需正式 plan 的合法范围须保留：新建控制仅经过明确创建 admission；已有独立记录以其已校验 envelope 证明 document/project，不能从 preference_save 顶层不存在的 projectId 猜 scope。
- import 有两个 scope：源 `source.documentId/projectId/changeId` 已验证保存记录只读，目标 document/project 使用 write lease；保留 digest、approvalId、expectedOrigin 约束。冻结源时不返回旧私密内容，冻结目标时不写目标，源内容绝不被删除。
- get/list/read 不 mkdir、不触发导入或修补。普通 delete 同样要 active write lease；统一删除执行器未来另走 deleting intent，不能拿普通 API 绕过 fence。

**RED→GREEN**：两个实际项目/同文档、跨文档 import、嵌套 preference alias；原项目排队冻结后拒绝；取消/冻结后 observation complete 与 preference publish 不发生；纯读没有控制或库文件时不创建；原 digest/CAS/symlink/配额测试保留。独立 scoped checkpoint。

## 单元 4：research 异步存储与 team ACL

这是两个可并行交付、可分别拒绝的子单元。

### 4A research

**所有权**：`presentation-research.ts`、`packages/project-store/src/presentation-research-store.ts`、新增 Shell `presentation-research-write-fences.test.ts` 与 `packages/project-store/tests/presentation-research-write-fences.test.ts`；保留 research/abandon/cleanup 既有测试。必要 lease 最小扩展由单元 1 helper owner 合并。

- 请求、parsed draft、scope、ledgerId/deleteId/expectedRevision 在 `acquireProjectLock` 前固定；build 的 begin/每次 evidence await/finish 与 catch finish 使用同一 lease。abandon/delete 的 unprotected check await 后也不能续租。
- 存储 `update()` 自己还 await 私有锁、load、mkdir、temp cleanup、open/write/fsync/rename。必须向它传播同一同步 guard；仅在 service 调 `store.finish()` 前检查不够。围住每个实际 mutation，并在最终 canonical state.json publish 紧前检查。若 publish 使用异步 rename，冻结可在排队后执行前进入；采用紧邻 guard 的同 JS turn 最终发布，或等价明确序列化，不声称跨进程 CAS。
- pre-plan build 必须仍可用。readonly legacy research 的真实 scope 在 load envelope 中；研究单元需要异步 scope proof，helper 新增 `capturePresentationProjectAsyncReadLease({store, scope, readExistingProject: async (scope) => actualBoundScope | undefined, signal?}): Promise<PresentationProjectReadLease>`，复用已有内部快照/检查：**先**同步固定 lifecycle/revision 或 absence，再 await 有界 namespace proof，再检查原快照并返回仅 assertCurrent 的 read lease。不得 await proof 后才 capture 最新版本，也不得借此获得 write 权限。空 history 不初始化控制；已有实际记录/删除回执校验自身 document/project 后才返回。
- 已发出的异步 I/O 不假装被取消；私有 staging 只清理本 invocation 确认拥有的临时文件。项目删除须 drain 此任务，canonical publish 在 freeze 后禁止；存储安全/校验和/原研究 tombstone 与 archive 语义不降级。

**RED→GREEN**：研究先于 plan 的真实 PC 用例；延迟 evidence success/failure freeze 后不 finish；暂停在 update 私有锁/load/temp-open/最终 publish 后冻结；source draft alias；delete unprotected await 后 freeze 无删除；只读未建控制字节一致；重开 running 可观察但不自动失败/重放；版本 1/2 archive 与删除回执原保证仍成立。

### 4B team

**所有权**：`presentation-team.ts`、既有 `presentation-team.test.ts` 与新增 `presentation-team-write-fences.test.ts`。

- 深拷贝 validated body 和 validated context/expectedIdentity。teamId 操作先授权实际 ledger，再从 ledger 取 project/document capture lease，不能信请求补出的 scope。
- create/publish 的 await 项目锁前 capture；所有 member/comment mutations 也走同一项目锁，不只 publish 加锁。锁后重新授权当前 ledger/ACL 并检查原 lifecycle，业务 expectedRevision 不刷新。
- 初始 create 的 store.write 和公共 commit 的 store.write 紧前 guard；纯 identity 不需要项目 lease，project_read/plan_read 有固定 read lease及最终返回检查。不可因纯读初始化生命周期或改 ACL。

**RED→GREEN**：跨实例 publish/member/comment 排队冻结；await 中 revoke actor 后不得写；真实 ledger scope 与请求/context alias；deleted 项目不重建 team snapshot；合法 owner/reviewer/viewer 原 ACL 与 comment 上限不缩减；pure reads 目录字节一致。4A/4B 分别 scoped checkpoint。

## 单元 5：delivery bundle 与项目 page backup

可分两个子模块 owner；service factory 接线归主 owner。

**文件**：`presentation-delivery-bundles.ts`、`presentation-page-backups.ts`、既有 `presentation-delivery-bundles.test.ts`、`presentation-page-backups.test.ts`、`presentation-page-backup-integration.test.ts`；分别新增 `presentation-delivery-bundle-write-fences.test.ts` 与 `presentation-page-backup-write-fences.test.ts`。

- 先固定请求、实际 production/lineage scope、原 lease，再进入外部项目锁与模块私有锁。只读使用 directory(create=false)，缺目录返回原合法空/不存在结果。
- delivery 现有 `metadata()` 会 cleanupTemporary；list、read、metadata 不得为“查看”清理 temp。把 cleanup/receivedBytes 持久修补移到明确写分支；chunk 的 exact-overlap 重试在 active 原 lease 下收敛 receipts，保持 lost-ACK 上传进度，不让只读伪造已保存回执。write 分支各 mkdir、temp/open/write、metadata save、publish、delete 都需原 guard及 drain 登记。
- page backup 保留已 frozen body、实际 child/parent lineage 和 hash。begin staging→publish、chunk 的 `await open/stat` 后实际 fd.write、finish ZIP 验证后的 metadata publish 分别检查；status/read 不 mkdir 或标 ready。
- 回执 ACK 丢失按真实 stored bytes/hash/metadata核对；冻结后只提供治理需要的有限观察，不盲目重发 write/finish/delete。不是删除确认本身。

**RED→GREEN**：begin/chunk/finish 在私有 lock、ZIP 解压或 fd-open await 后 freeze，无晚 canonical 发布；成功/失败 ACK 不伪成功；重开原 exact-overlap 客户端继续进度；read/list/metadata 缺控制、缺目录、已有临时文件时字节不变；同文档两项目不互相阻断或改包；page parent/hash/配额/ZIP/路径安全断言不削减。两个子模块分别 scoped checkpoint。

## 单元 6：主接线、全范围证明与删除执行器前置合同

主 service owner 串行接入各模块固定 lease factory 和任务登记，其他 owners不再编辑 service。共享作用域服务仍使用原 authorization；只有未来显式项目引用动作受本项目 fence，文档级 blob/品牌不得被猜测绑定。

验收使用合成临时 PC 两个 service 实例：

1. 枚举所有上述写入口，在 deleting/deleted 与旧 revision 上均无新普通项目目录/正文/事件；创建控制记录只在明确创建入口且默认保留，不复活 ID。
2. foreground 编译已持主锁时：先同步保存 deleting intent，锁外 abort+drain，不能中断则等真实 compiler 结束；旧请求 guard 拒绝 complete，释放主锁，再供删除执行器取得锁。排队请求也 settle，不能锁内等它们。
3. production pause/current-page、全部历史 read byte-equality、600 页/原资产协议/ACL/研究恢复用例保持；测试仅扩验新的具体风险，不重复无关昂贵母版压力。
4. 原异常/取消返回不泄漏正文；治理观察与最小 audit 可以返回状态/计数，不能经冻结 status/export 绕出旧私密内容。
5. 每个资源执行器将来使用 deletionId + 固定删除 revision/原真实归属，不调用已经拒绝 deleting 的普通 write API来绕过栅栏。shared/unproven资源保留并返回 partial。

此合同通过前，不接实际自动保留期清理。发布只合并已独立审查的 scoped commits；迁移保留 legacy 正文与控制 absence，不能批量扫全库补控制或绑定共享归属。回退不删除已写的 lifecycle/tombstone，也不能退到无 fence 的版本后允许项目重写；可暂时关闭受影响写能力并保留只读治理/恢复。最终验证包括 fresh Shell/store types、lint、受影响全套与实际服务 races；真实宿主/专业门禁和整体完成度仍按原方案另行评估。
