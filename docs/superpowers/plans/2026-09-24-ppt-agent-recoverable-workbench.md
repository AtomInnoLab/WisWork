# O1 可恢复生产工作台批量实施

基线4821f1e8。用户要求一次完成多项关联任务，避免只补小项。依据原方案O1、§9职责划分；继续现有隔离分支，不改原方案，不实机验收。使用高保证/子代理实现/独立审查/TDD；复用现有逐页生产与产物导入。

## 一次交付范围

PC后台页生产不绑定请求signal；持久任务控制与生产事件；Taskpane恢复、暂停/继续/取消和事件视图；保存已完成单页与完整成果导入准备入口。PC进程重启后标明中断，由用户继续，成功页不重复编译；不自动重试无限失败、不写Office、不将编译或时间线当作QA通过。后台任务是冻结SlideIR的页编译，不声称自动续跑研究/LLM推理/Office写入。

## 接口

新增PC operations: production_job_start/status/pause/resume/cancel，精确四字段{operation,documentId,projectId,requestId}。响应{job:PresentationProductionJob|null,production:既有production summary加inputDigest/planDigest（仅job操作）}，统一最大256KiB。status不存在job时null，但production必须存在并验证文档绑定。start从无job启动；resume仅paused/interrupted/failed；running重复start/resume幂等不另启动；cancelled终态不重新启动；completed幂等。暂停/取消在正在编译页完成后生效，保留该页，不提前宣称停止；queued操作接纳后立即返回。取消不删除任何成果。

浏览器安全共享契约新放packages/project-store/src/presentation-job.ts，以./presentation-job子路径导出，供Store/PC/plugin使用，不引入Node或引擎依赖。Job字段version1,projectId,documentId,requestId,inputDigest,planDigest,planRevision,revision,state,events。state: running/pausing/paused/cancelling/cancelled/interrupted/completed/failed。events:{sequence,createdAt,type,pageId?,attempt?,error?}；type: run.started/pause_requested/paused/cancel_requested/cancelled/interrupted/completed/failed, page.started/compiled/failed。错误枚举沿用编译错误加invalid_state（仅run.failed）。每次追加revision+1，保留最近128事件，序列连续且最后序号=revision；UI明确更早历史已截断，不能伪装完整审计。固定UTC毫秒时间；事件字段按类型严格验证、无任意文本、未知字段拒绝。

Store方法 productionJob(projectId,documentId,requestId):Job|undefined；appendProductionJobEvent(projectId,documentId,requestId,expectedRevision,event):Job，event不含sequence/createdAt。新job预期revision0。原子文件带摘要、绑定production输入/计划摘要；page事件核对真实保存状态/attempt，完成事件核对所有页compiled，状态转换严格且CAS防陈旧。文件损坏拒绝，不覆盖；独立文件无旧记录迁移。

后台runner进程共享注册表key=resolve(userDataPath)+projectId，同项目只运行一个job（异request返回busy），跨文档/项目独立。通过既有service项目锁协调启动与foreground生产run，后台自身不持有跨页长锁以允许状态/暂停读取；foreground production_run在后台任务运行时返回busy。每页更新production receipt后再记对应job事件；若崩溃在两者之间，以真实页receipt为权威，重开标记中断并保留成果。worker有独立controller不接受client abort，pause/cancel为页边界flag；启动请求在接纳前取消则无副作用。后台Promise必须捕获所有失败并释放注册表。无活跃worker却有running→interrupted、pausing→paused、cancelling→cancelled，读取可完成该恢复转换。同一进程多service实例共享worker不得误判中断。用户明示resume继续，成功页跳过，失败页每次run最多尝试一次。

## 分工及验收

A Store+共享契约(agent): packages/project-store/src/presentation-job.ts、presentation-store.ts、index.ts、package.json及对应tests。TDD状态机/原子持久/校验/重启/坏文件/跨doc/CAS/128事件回放/成功页绑定，限定提交。
B PC runner(agent): apps/shell/src/main/presentation-jobs.ts、presentation-service.ts、presentation-production.ts及对应tests。handler新增可选onPage(record,page)/shouldStop()，在每页写入前后发事件、页边界暂停，旧foreground行为保持；新operations与busy准确分类。TDD断开signal仍继续、暂停/取消页边界、重启、重复启动、并发项目、失败重试、状态期间可读和锁竞态，限定提交。
C 工作台(agent): 新插件skills/powerpoint/presentation-jobs.ts，presentation-project.ts，agent/presentation-project-card.tsx，styles.css及相关tests。共享parser校验响应身份/摘要一致性/体积/过期guard；新独立skill工具映射五operations；controller从status拉job并管理操作、轮询时机和清理，不允许旧文档数据发布。UI显示job状态/最近事件/下一动作、暂停继续取消，普通取消等待不等于取消PC任务；保留已有功能和旧PC fallback，修正页替换已接入却仍称未接入的文案。controller公开downloadProductionPage(pageId)、prepareProduction()用于已编译页/全完成；通过existing工具执行，无宿主写入。限定提交。
Root: host-runtime工具及controller路由/清理/文件变化通知接线，跨层集成（使用真实PptxGenJS、真实Store，mock仅宿主）、阶段文档与总体重新评估。独立全diff审查；最后全仓test/typecheck/lint/format/licenses、Addin/Shell build。

## 回滚与边界

无新依赖，无旧数据迁移，无自动Office写入。旧PC不支持job时清楚提示升级，仍可使用现有前台页编译；旧插件可忽略新增job文件，不删除记录降级。任务数量受既有production上限约束，后台并发以现有项目隔离为基础；不声称跨PC进程分布式锁。事件日志是生产阶段恢复视图，不冒称包含研究或宿主QA事件。整体42%基线，完成后按同一九模块档位评估，不预设必须上涨。实机0/20继续暂缓。不合并/推送/部署。

## 审查修正

项目status必须优先暴露运行中request；增加生产任务清单和选择器以恢复旧paused/interrupted/failed任务。Store productionHistory复用校验读取，返回按sequence倒序的clone。status增加productionTasks精简字段（requestId/sequence/planRevision/status/compiledCount/total/jobState可选），不重复全部页面，最多32。Taskpane明确选择的请求在同文档刷新中保持，切文档清理。单页派生修订隐藏整套导入准备入口，保留单页保存与既有替换工具。插件声明对现有本地project-store包的直接依赖，仅引入browser-safe子路径，无外部新依赖。

## 完成记录

四个交付单元及审查修正均已完成；全仓测试6,342 Vitest +11 Node +53 Rust通过（6跳过）、类型/lint/许可证/格式与双端构建通过。独立审查复审无阻塞项。O1按主要工程工作台链50%→75%，整体42%→44%，实机0/20继续暂缓。完整研究到宿主的AgentRun恢复仍列缺口。详见阶段报告 `docs/product/wiswork-ppt-agent-recoverable-workbench-progress-2026-09-24.md`。保留实现分支与工作区。
