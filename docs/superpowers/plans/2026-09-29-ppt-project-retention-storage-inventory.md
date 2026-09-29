# 项目保留期与统一删除：存储核对及实施边界

按原方案§13.1核对，基线93df802f。本文为下一实施单元的输入，尚未实现保留期或统一删除。

| 存储                                                                                | 归属                              | 删除需要解决的问题                                                                                                  |
| ----------------------------------------------------------------------------------- | --------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| projects/presentations/<projectHash>                                                | 项目，内部绑定文档                | 计划/修订/生产/编译产物/来源审查/任务/反馈/资产事件；统一停止worker与禁止新写                                       |
| presentation-research/<documentHash>/<projectHash>                                  | 文档+项目                         | 正式记录、归档及删除回执，不能遗漏已有清理状态                                                                      |
| presentation-delivery-bundles/<documentHash>/<projectHash>                          | 文档+项目                         | 完整包、上传中暂存及清单，和项目写锁协调                                                                            |
| presentation-page-backups/<projectHash>                                             | 项目                              | revision原页备份及暂存，删除前处理运行中的页修改                                                                    |
| presentation-preferences / presentation-comments / presentation-manual-observations | 文档+项目组合hash文件             | 包含本项目内容，需精确路径及内部归属核验                                                                            |
| presentation-teams                                                                  | owner+document+project生成team ID | ACL和已发布内容，不能仅删除私有计划                                                                                 |
| presentation-attachments/<documentHash>                                             | 仅文档                            | 目前无项目独占归属；建立显式引用/所有权，不能误删同文档其它项目资料                                                 |
| presentation-acquisition-history/<documentHash>.json                                | 仅文档                            | 含远程来源和附件结果；需按引用处理，不能把来源审查当保密审计                                                        |
| presentation-existing-page-backups/<documentHash>及.released                        | 仅文档                            | 修改事务仍引用原页；需项目关联与未决保护，不能为删一个项目清空整文档                                                |
| presentation-master-backups/<documentHash>/<changeHash>                             | 仅文档+修改事务                   | snapshot、全部原页、图片与逐项回执证明；当前不含项目独占归属，未决/恢复事务和共享引用须保护，不可按当前项目猜测删除 |
| presentation-package-backups/<documentHash>/<changeHash>                            | 仅文档+XML修改事务                | 原始/准备包、完整页序摘要、阶段证明及历史复核；无项目独占归属，未决与恢复引用须保护，不能按当前项目猜测删除         |
| presentation-brand-kits                                                             | 用户全局共享                      | 本项目引用解除；无授权不删除全局品牌资源                                                                            |

## 实施顺序

1. 持久项目生命周期(policy/revision/deletionId/state)，全入口统一在项目锁下校验，worker与长下载完成前二次校验；删除中的项目不能产生新记录。
2. 建立项目资源引用和历史存量归属核对，文档共享资料仅移除明确项目引用；未证明独占的资源不得删除。明确给出未清理项，不能将partial宣称complete。
3. 用户可见预览与确认后保存删除intent，分资源持久回执，失败可重开按实际剩余状态继续；所有文件操作拒绝symlink和目录越界。保留期到期走同一删除事务。
4. 持久、可导出的最小审计记录：动作/时间/本项目匿名标识/结果/回执摘要，排除正文、附件、原始URL、完整documentId。审计保留策略与项目内容保留策略分别明确。
5. 合成PC临时目录集成验证多项目共享、并发上传/worker、进程重开、断点删除、ACL、预算、路径攻击；不得清理真实用户数据。

该范围完成后仍需原方案宿主/专业门禁证据，不能用测试数量推升整体百分比。

## 母版交付后的核对（2026-09-29）

当前通用项目删除只转入 `.trash` 并重置文件项目映射，没有联动全部 PresentationStore。实施需要统一生命周期锁及版本复核：先阻止新写，停止并等待 worker，再逐资源写入持久删除意图和回执。文档级历史存量没有项目独占证明时保留，并显示 partial/blocked 清理结果。

审计保存随机匿名项目标识、动作、时间、结果及最小回执状态；不写正文、附件、URL、路径、完整文档 ID、原始内容摘要或原始错误。审计保留期与项目内容策略分别配置。首次集成只用临时目录与合成数据，不删除用户真实内容。本核对仍不是保留期/删除功能完成。

## XML 母版恢复并行核对（2026-09-29）

`presentation-package-backups` 现在同时承载页面/图表 `package_xml` 与母版 `master_xml`。后者还需留存未证明写入的当前受影响页包与差异证明；这些记录仍只有文档/修改事务归属，不能推测项目独占归属。删除预览必须把无法证明独占的原包、当前包和回执链列为保留/未清理项。

已核对 `presentation-service.ts` 的项目锁是进程内按 `userDataPath + projectId` 串行化，并非持久生命周期栅栏；`presentation-production.ts` 的异步 worker 需要在写入结果前检查持久生命周期版本。统一删除需要先发出取消/禁止新写，再等待项目锁与 worker 停止，不能在等待同一长任务锁时才尝试取消。generic `store.ts#deleteProject` 只移入 `.trash` 并重置 fileMap，不能当作 PresentationStore 完全删除或隐私删除成功。

## 实施单元与当前边界（2026-09-29）

- 生命周期记录放在独立 `presentation-project-lifecycles/<projectHash>`，核对已有项目绑定，不改变正文目录和共享资料。持久版本 CAS、保留策略、删除 intent、逐资源结果和最小匿名审计；当前存储事务按 PC 单主进程模型串行，不宣称跨 OS 进程原子 CAS。
- 默认内容/审计保留期为 null（保存直到用户显式配置），不能无配置自动删除。策略写入不等于定时删除执行器已上线。
- 只读资源清单核对每个命名空间的真实 document/project 归属。独占项列候选，共享或未知项列保留；候选仍须在持久冻结、后台任务停止和实际删除前再次核对。
- 删除必须先持久进入 deleting 并禁止新写，发出 worker 取消，再等已运行工作结束；不能等同一项目锁之后才取消持锁 worker。最终写回前核生命周期版本，不能在删除后回写旧产物。
- 全资源移除/解除项目引用才可进入 deleted；任何保留、失败或无法证明独占的资料均明确保留 partial。断点重开使用现有 intent 与结果，不创建第二次删除或伪造完成。

当前生命周期与资源清单基础已提交 4fd13074；后台 guard/停止等待接口和四项 PC 元数据路由已完成。尚未接通全部写入口、用户确认界面或实际统一删除。只用合成临时目录测试，不清理真实用户数据。

## 服务与后台栅栏接线

以下保留实施前入口核对与后续接线约束。后台 guard 与停止等待已实现，实际服务持久 lease 接线仍在进行，不代表统一删除已完成。

### 固定请求身份与持久版本

使用 `PresentationLifecycleStore.assertActive({ projectId, documentId }, expectedRevision)`：请求获准时固定 scope/revision，等待项目锁后、每次异步返回后和每次同步落盘紧前复核。版本只能由新的请求或明确恢复操作重新取得，不能在旧任务每次检查时悄悄改用最新版本。`presentation-service.ts` 当前 `request = body` 和 `presentation-jobs.ts` 的延迟调度仍保留可变输入引用，应在首次 await 前深拷贝已验证、受预算约束的请求，并固定项目、文档、requestId 和锁 key。

### 服务入口和最终写入点

- `presentation-service.ts` 的主项目锁在 `acquireProjectLock`（约第58行），普通项目操作约第983行取得锁。锁是主进程内串行机制，不能代替持久生命周期 fence。
- 需要在实际调用紧前复核：`recordProductionFeedback`、`acceptPlan`、`setPlanPageLock`、`savePlan`、`begin` 和 `complete`。尤其 `compile` 持有项目锁跨编译 await，最后 `store.complete`（约第1606行）只有 signal 检查，没有持久版本复核。
- 主锁之前提前返回的 comment add/resolve、manual-observation begin/complete/delete、preference save/import/observation，以及 team/research/delivery-bundle 分支都需独立接线。异步子服务必须在下载、解析或研究结束后的实际写入处使用同一 fence；仅在外层服务进入时检查不足。
- `production_job_status` 并非纯只读：无匹配 worker 时会追加 interrupted/paused/cancelled 事件。冻结后不得借状态查询创建普通项目事件；可只返回观察结果，或在 active 原版本下执行恢复事件。
- PDF 导出绕过主项目锁且跨渲染 await。尽管不持久写文件，也必须在读取前绑定版本、返回最终字节前复核，防止项目冻结或删除后仍返回旧私密内容。
- 文档级附件、page/master/package 备份和品牌资料没有项目独占证明时仍是共享或未知归属，不能根据请求所在文档自动绑定为某个项目资源；项目引用的新增则须受项目 fence 约束。

### Worker 注册、取消和逐次写入

`presentation-jobs.ts` 当前进程级 workers 只存 requestId，AbortController 在启动分支局部创建，取消接口无法直接 abort。后台 worker 在受理响应之后运行，本身不持上述项目锁。最小注册项应持有不可变 `{ scope, revision, requestId, controller, settled }`，并提供跨 service 实例可用的按项目取消、等待结束接口。foreground 长操作也需可取消的登记，或明确等其实际结束且依靠持久 fence 拒绝晚到写入；不能声称编译器已支持中断。

向 `handlePresentationProduction` 注入同步 `assertWritable()`：检查原 lifecycle revision 和取消 signal，围住 `saveClaimReview`、`deriveProduction`、`beginProduction`、所有 `updateProductionPage` 和 `appendProductionAsset`。资产失败的 rejected 事件、页面失败的 failed 状态同样属于写入，不能在取消/删除后从 catch 补写。jobs 的 append/onPage/完成/catch 事件也需要原版本 fence。

保留既有 pause 语义：`shouldStop` 继续停止新页面调度，pausing 允许已经运行的页面完成；不能直接将 shouldStop 作为每次写入拒绝条件。cancel 才 abort controller，并在每次存储前阻止晚到结果。正常 active 项目的终结 job 事件可以使用生命周期 fence，但 deleting 项目不得写普通失败/取消内容事件；删除取消意图和进展属于独立生命周期账本。

### 删除顺序与恢复

先持久进入 deleting 并保存删除/取消意图，然后在项目长锁之外发出取消、等待 worker settled，再取得项目锁核对和清理。不能先等持锁的编译/下载任务，再尝试取消它，也不能让删除在锁内等待一个还需要同一把锁才能退出的 worker。不能中断的编译器仍须等待实际结束，晚到产物不得落盘。每项清理前复核 deletionId、版本和资源真实归属；共享、未知或失败资源保留 partial，不宣称完全删除。

### 必须覆盖的合成并发用例

1. foreground compile 持锁时开始删除：先冻结成功，编译晚返回无法 complete，等待和释放无死锁。
2. background asset-fetch/compile 停在 await 时取消或删除：成功、失败、rejected、failed、job catch 路径均不得晚写。
3. 请求在冻结前排队、冻结后取得项目锁：拒绝旧版本，不创建新目录或记录。
4. pause 允许当前页完成并停止新调度；cancel 阻止当前页晚写，两个控制的原能力保持。
5. 两个 service 实例共同观察/取消同一进程级 worker：取消命中正确 scope/requestId，drain 不依赖单个实例局部 controller。
6. 进程重开保留 tombstone：不得恢复 worker 或新建内容；冻结状态查询不追加伪只读恢复事件。
7. PDF 渲染迟返回、调用方修改请求别名、active 项目 revision 更新：旧版本结果和项目身份变化均被拒绝。

这些用例只使用临时目录和合成项目；本节不授权删除真实用户内容，也不声称跨 OS 进程原子 CAS 或未来 PC 容量预留。
