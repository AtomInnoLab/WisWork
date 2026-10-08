# 演示项目删除入口门禁核验（2026-09-30）

## 范围与结论

核验基线：`a4c664e1`（Unit4 实际 Service 接线）。范围仅为 `createPresentationService` 当前公开的演示项目入口，以及它调用的项目写租约、Work、worker 和 namespace factory；不代表全仓安全审计、实际 Office 验收或删除执行器验收。

当前已核对的项目独占正文写入口均接入固定生命周期版本检查。未发现仍经这些 Service 入口绕过 `deleting/deleted` 的普通正文写分支。**这只是公开删除路由的前置条件，不授权删除，不表示删除路由已可开放。** 后续删除执行器、资源归属复核、未决事务保护及 UI 确认仍需独立门禁。

核验 source：`apps/shell/src/main/presentation-service.ts`；相关实现为 `presentation-project-write-lease.ts`、`presentation-project-work.ts`、`presentation-jobs.ts` 与下表 factory。直接调用低层库时可选 guard 的兼容行为，不等于实际 Service 路由缺省不接 guard。

## 入口矩阵

“固定”指首次业务 await 前捕获 lifecycle revision 或 absence；await 后只检查原值，不续租。`Work` 的结束必须来自实际 finally，取消通知本身不是已停止。所有 namespace 仍保留原 schema、业务 CAS、身份、摘要、配额和重试回执。

| 实际入口                                                                                                                                                          | 真实项目归属 / 首次受理                                                                                                                 | 写栅栏与锁                                                                                                                                                                                 | Work / 返回门禁                                                                                                             |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| `compile`、`save_plan`                                                                                                                                            | 校验 deck/plan project；既有 `project.json`；合法新建才显式 creation lease，默认保留策略                                                | 主项目锁前固定 lease；编译、研究、资产 await 后检查；begin/complete/save 紧前检查                                                                                                          | Service 独立 foreground token，实际 finally 释放；冻结后成功或失败都不能补正文                                              |
| `resume`、`accept_plan`、`set_plan_page_lock`、`audit_sources`                                                                                                    | 已存在 metadata scope；unknown 保持原 not_found                                                                                         | 原版本 write lease → 共用主项目锁；source audit begin / success / failed finish 均有最终 guard                                                                                             | Service foreground；取消/冻结不补普通 failed 回执                                                                           |
| `get`、`status`、`get_plan`、`read_import_source`、`read_source_audit`、`export_pdf`                                                                              | 已有实际 metadata，合法 legacy absence 保持 absence                                                                                     | read lease，显式无 write 权限；普通读走主锁；PDF 每次 render/load/copy/save await 和 catch/最终字节返回复核                                                                                | Service foreground；旧版本不能返回私密正文或 PDF；只读不创建 lifecycle                                                      |
| `production_begin`、`production_rebuild_page`、`production_run`、`production_record_claim_review`、`production_record_issue_action`、`production_feedback_record` | 实际 Store 项目 / frozen production；请求、plan、页与业务 CAS 原校验                                                                    | 固定 write lease → 主项目锁；Production/DeliveryReport 的实际 Store 更新紧前 guard                                                                                                         | Service foreground；await 后与返回前复核；未知项目不凭请求创造正文                                                          |
| `production_status/page/content_check/page_reviews/claim_evidence/read_claim_review/delivery_report/feedback_read/feedback_compare`（均带 `production_` 前缀）    | 实际 metadata / production 版本                                                                                                         | read lease；absence 不初始化；误用写 guard 为 access_denied；共用主锁                                                                                                                      | Service foreground；结果返回前检查原版本，失败不能变成空“成功”                                                              |
| `production_job_start/status/pause/resume/cancel`                                                                                                                 | 真实 frozen production；status 可能补恢复事件，故按写受理                                                                               | 受理时固定 write lease；主锁保护受理 / 事件；后台每次页写与 job event 均守原 lease                                                                                                         | 前台受理 token 与后台 worker 分离；后台自有 controller，不继承已结束客户端 signal；`stopPresentationWorkers` 等实际 settled |
| `comment_add/resolve`、`comment_list`                                                                                                                             | 必须真实 saved plan；list read lease；合法库 final write 前才受理 legacy write control                                                  | 同步库无业务 await，不需持长主锁；输入 / CAS / not_found 在 control 创建前拒；temp/rename 最终固定 guard                                                                                   | Service foreground；每请求独立 token，finally 释放                                                                          |
| `manual_observation_begin/complete/delete`、`preference_save_observation`；观察 get/list                                                                          | begin 严格 shape/ID 后显式 creation；其余实际 observation envelope scope；before/after digest CAS 保留                                  | 写库同步最终 guard；读取只有 read lease，空查询无控制/正文创建；保存偏好仍核 origin                                                                                                        | Service foreground；实际记录绑定，不将 observation 的自报来源冒充作者认证                                                   |
| `preference_save/delete/import`、`preference_get/list`                                                                                                            | save 的 scope 来自 shared parsed item；get/list/delete 来自实际 namespace；import 源实际记录 read lease + 目标独立 write/creation lease | 同步最终 source/target guard；原 digest/approval/expectedOrigin/CAS 不降级；空 delete 返回 false 不授写                                                                                    | Service foreground；两个项目不混成同一租约；源冻结也不能返回旧私密偏好                                                      |
| `research_build`；`research_abandon/delete`；`research_list/latest/read/delete_status`                                                                            | build 已严格 parse draft 后显式 creation，可合法先于 plan；历史由真实 research summary scope 或项目 metadata 证明                       | build 首次 await 前固定 write lease；历史异步 proof 前固定 read revision/absence，proof 后检查原值；写升级前仍核同一原版本；共用主锁 + Store 私有锁；begin/finish/update/delete 最终 guard | factory 自有 Work，finally；证据 await 成功/失败不能冻结后 finish；纯 empty/legacy 查询不 mkdir/control，不作为写授权       |
| `team_project_create/plan_publish/member_set/member_revoke/comment_add/comment_resolve`；`team_project_read/plan_read`                                            | create/plan_read 的真实 saved plan；其他由已授权实际 ledger 取 doc/project，不能信请求虚构 scope                                        | fixed lease 在共用主锁前；锁后重新 ACL / ledger / CAS；最终 Store.write guard；temp wx 未创建则不删外来文件                                                                                | factory 自有 Work，finally；owner/reviewer/viewer 原规则保留；`team_identity` 无项目 scope，无项目 Work/lease               |
| `delivery_bundle_begin/chunk/finish/delete`；`delivery_bundle_list/metadata/read`                                                                                 | factory 先读取实际 production，冻结 accepted scope / manifest 绑定                                                                      | fixed lease → 共用主项目锁 → 模块私有队列；写 mkdir/open/bytes/receipt/publish/delete guard；只读不 cleanup/repair                                                                         | factory 自有 Work；锁等待也登记并能 drain；lost ACK 仅由明确 chunk/finish 重试收敛，不由读请求改回执                        |
| `page_backup_begin/chunk/finish`；`page_backup_status/read`                                                                                                       | 真实 child/parent production lineage 与页；Service 注入 metadata scope lease                                                            | fixed lease 在模块私有项目队列前；fd open/stat/write、ZIP 验证 / metadata publish 最终 guard；只读不标 ready                                                                               | factory 自有 Work，finally；**目前使用 namespace 私有锁，不使用 Service 共用主锁**，不能仅凭取主锁认为此任务已结束          |

纯查询的 matching document namespace 确实为空时，部分旧 API 继续返回空 / not_found；这种兼容分支不证明项目归属，不创建 control/body，不授写。真实资料、显式创建、不同 document 重用已有非空项目 ID 仍受全局生命周期绑定拒绝。

## 治理 metadata 不得复活项目

`project_lifecycle_*` 为同步治理控制操作，不是普通内容写，采用 LifecycleStore 的同步 transaction / CAS，不等待异步项目主锁。

- `set_policy`：必须精确 expectedRevision，随后 `assertActive`；deleting/deleted 拒绝，旧 revision 为 revision_conflict。政策更新会增加版本，旧内容请求不能借它续租。
- `initialize`：缺控制时只新建控制 metadata，不创造项目正文；已存在时**幂等返回原 record**，包括原 deleting/deleted 状态，既不重写、不改状态，也不清除 tombstone。因此它不一定返回错误，不能将这个响应当作“项目已 active”或写授权。
- read/export_audit：不授正文写权限；匿名 audit/control 独立保留，不随项目正文移除。错误 document scope 拒绝。取消或 malformed 字段不产生控制更新。

## 文档共享 / 全局资源保留边界

以下 Service 分支不因为请求处于某项目便获得项目独占归属；不能将它们直接加入项目正文删除清单：

| namespace / 入口                                                          | 删除边界                                                                                            |
| ------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `existing_page_backup_*` 原页包、render、release/abandon 与未绑定 staging | 文档 change/backup 级恢复证据。项目无独占证明时 retained/unproven；不能猜测将释放当删除项目全部原包 |
| `master_backup_*`、`package_backup_*`                                     | 文档/savepoint 级 blobs 与 manifests；保留实际引用链，不从当前 projectId 推断独占                   |
| `attachment_*`、acquisition history、原文件/提取资产/许可记录             | 文档共享；只移除经过实际证明的项目引用，不删共同 blob 或跨项目来源                                  |
| `brand_kit_*`                                                             | 全局品牌库；仅真实项目引用可作为引用项，品牌本体 retained                                           |
| 未绑定 page / existing staging、未知或损坏归属                            | retained/unproven 或明确拒绝扫描；不可当作已清理，也不能靠猜测归属删除                              |
| lifecycle control / audit / tombstone                                     | 独立治理 namespace，按独立 audit policy 保留；正文删除不得清除防复活状态                            |

Inventory `candidate` 仅表示已证明项目独占的候选，仍不是删除许可；`complete:true` 仅表示约定有界 namespace 扫描完整，不表示全部用户资料已覆盖。

## 公开删除路由必须继续满足的门禁

1. 显式用户确认冻结的 exact inventory / scope / 原 revision，持久 CAS 保存 deleting intent；不能先删文件再登记。
2. 冻结之后，**在共用主锁外**调用 `stopPresentationProjectWork`：精确 root/project/document，abort 全部 foreground token 和匹配后台 worker，并等待每个真实 settled。PageBackup 私有锁和后台 compiler 尤其不能由“已取主锁”替代 drain。
3. 无 signal 的编译或已经发出的 I/O 只能等真实结束；不要宣称即时中断。请求最后 finally 之前不能开始清理。
4. drain 完成才获取共用主项目锁，再核 deletingId / revision / 实际归属与 namespace/文件身份。executor 使用独立 deletion guard；不得调用已经拒绝 deleting 的普通 write API 绕过栅栏。
5. 逐资源保存固定 receipt；unknown/共享/未决事务保持 retained 或 partial，不返回“全删完成”。取消、崩溃、ACK 丢失只能按真实持久 intent / receipt 重开；不能盲删第二次。
6. 生命周期初始化幂等响应、政策更新、不同文档、Worker 重开、纯读修补、共享备份均不得清除冻结或复活原项目 ID。

这些门禁尚不代表删除执行器 Unit2 独立审查通过，也不开放真实删除 route。本核验没有执行任何用户资料删除。

## 验证证据

- Unit3：实际同步 temp 写后冻结、source/target import、legacy malformed comment 受理；六套 42/42，`/tmp/b-sync-comment-green.log`，root 独立审查后提交 `e89ba34f`。
- Unit4：真实 Service 四 factory deleting、研究 pre-plan/纯 empty read、锁等待期间冻结、Team 真实 drain、外来 temp 保留；九套 79/79，`/tmp/b-factory-regression-final.log`；独立 29 与 root 32 通过后提交 `a4c664e1`。
- 本轮 existing gate 验证：ordinary/PDF、Production lease、同步库、factory、Work、Lifecycle、PageBackup 共 10 文件 / 106 项通过，`/tmp/b-deletion-entry-gates.log`。没有重跑 600 压力测试，没有执行删除；未将删除执行器未验收项混入本矩阵。
- 本轮仅新增本 Markdown 文档，生产 source 未改变；沿用已冻结 Unit4 Shell tsc / scoped lint 的 exit 0 证据，不将文档编辑声称为新增 runtime 类型修复。
