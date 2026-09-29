# 研究中断记录的显式结束与恢复

基线bc25b166。依据原方案研究流程、持久项目、失败恢复和容量治理；沿用原已批准实施范围及隔离分支。高保证三单元、交叉独立审查；只修改源码并测试合成临时目录，不对真实资料执行结束/清理，不部署。

## 已确认缺口与设计

ResearchStore.begin同ID幂等返回原running，PC进程在begin后退出时不能finish；running禁止cleanup，容量可能永久占用。提供用户明确确认的结束，不自动重跑或猜测它已中断。PC和Store两层既有锁内只把精确running记录转failed/error:aborted，保留原draft/id/sequence/digest；不认证主张、不清空附件、不删除归档。实际活跃build持有同project lock，结束操作等待后若原build已完成则拒绝改写。

## 固定公共接口与兼容

Store新增 abandon(documentId,projectId,expectedRevision,ledgerId,expectedDraftDigest,signal?) → Promise<PresentationResearchRecord>。expectedRevision非负safeint，ledgerId≤128/digestSHA。读真实checksum/SHA、精确doc/project/ledger/digest核对；同ID已failed/erroraborted则幂等返回实际原Record（即使CAS已旧），不改revision/finishedAt/partial sources；completed或其它failed→record_not_running，不更改。running需要CAS匹配才结束，revision+1，finishedAt不早于原startedAt/全状态时间；failed sources省略（running原无证据）。beforecommit signal拒绝；latecommit abort不撤销，原record读取恢复。V1/V2存储形状和History版本不变，无新墓碑/操作回执，恢复读回的是原记录终态，不声称证明哪个请求完成了它。

PC新增research_abandon精确{operation,documentId,projectId,ledgerId,expectedDraftDigest,expectedRevision}→Record，仅现有pairedroute。既有research_read为状态查询，无新status endpoint；read只能根据实际Record判断running/failed/completed，不凭recent32历史缺失推断已结束。所有research操作仍用统一全局root\0project锁；abandon不删除所以无需计划引用扫描，但已完成研究必须拒，后续cleanup引用保护不变。所有输入大小/真实身份/safe error增加record_not_running。没有Agent abandon/mutation tool。activebuild/abandon并发必须实测保存结果不会被覆盖；processrestart orphanrunning实际可结束再cleanup、新ledger可build，原ledger永不重跑。

research_capabilities新增可选includeRecovery:true，必须同时includeCleanup:true；两者响应精确{version:1,available:true,cleanupAvailable:true,recoveryAvailable:true,historyVersions:[1,2]}。只有includeCleanup:true仍原旧cleanup exact；无flags仍旧legacy exact。Office纯helper readPresentationResearchCapabilities(request,options?:{includeRecovery?:true})：默认保留原请求与返回形状；optin先request cleanup+recovery，遇presentation_upgrade_required只回退cleanup，若再次upgrade回退legacy，最多3次。网络/响应错误不回退。只有optin接收5field新exact，旧两种响应合法且不显示recovery；返回旧fields+可选recoveryAvailable:true，false字段不新增。不擅改默认Agent capability。

## 单元与ownership

A remote_acquisition_pc：packages/project-store/src/presentation-research-store.ts +新tests/presentation-research-abandon.test.ts；Office skills/powerpoint/presentation-research-capabilities.ts +capabilities tests。按以上exact API/schema与旧字段保持；RED→GREEN actualV1/V2/restart/CAS/digest/idempotent/其它terminal拒/取消beforecommit/checksum/旧request三段fallback。只自己的files，不碰B/C/rootresearchSkill。后独立review B/root。

B addin_build_version：apps/shell/src/main/presentation-research.ts + presentation-service.ts safeerror; 新tests/presentation-research-abandon.test.ts。不改Store/cap/UI/rootfixture。capflagstrict/default旧shape；abandon同锁、新operation严格字段与safeerrors。用actual store.begin合成崩溃running，再service重开结束、read/cleanup/build不同ID，原ID build不得replay。真实 held source read/build与abandon顺序测试（可delay attachments原read入口，不mock最终service/record）。RED→GREEN；后独立review A/C。

C remote_acquisition_office：agent/presentation-research.ts controller、research-card.tsx、new recovery-storage.ts、host-runtime.ts hooks、App.tsx glue与相应tests。controller接口 abandonRecord(ledgerId,draftDigest)/retryAbandon()/checkAbandonStatus()；snapshot recoveryAvailable?/abandonAttempt?/abandonRecord?；phase增加abandoning/checkingAbandon。有限attempt精确{documentId,projectId,ledgerId,sequence,draftDigest,expectedRevision}，doc-scoped localStorage单独key，只metadata不存原文。Options readAbandonAttempt?(doc):unknown/writeAbandonAttempt?(doc,attempt|undefined):void，Runtimepresentation readResearchAbandonAttempt/writeResearchAbandonAttempt，App wiring同现有deletepersistence。持久身份保存失败不得发送；重开只读research_read（不auto abandon），显式retry原CAS，不自动重算或换ID。首次abandon仅summary.running精确digest/doc与CAS；receivedRecord必须doc/project/id/seq/digest和actualSHA吻合（root导出verifyPresentationResearchRecord(value) from researchSkill，C使用）；currentdoc/epoch/abort守卫在await前后，ACK和readstatus端到端。原记录completed或任意failed读取表示“记录已结束”而非声称本次操作取消成功；running仍pending。明确revision_conflict/record_not_running未提交可安全释放attempt并refresh，网络/invalid/cancel/abort模糊保留、只读恢复。record_deleted只明确提示原档已清理，不声称结束证明；not_found不凭空推断终态。storageclear throw安全publish/保留identity、不unhandled。不要让pending失败永久堵死：可明确forget pending仅清本机恢复metadata（不修改PC），产品提示丢失核对身份；或对安全终态足够释放，具体选择写tests。

Card：运行态记录显示“缺少结束回执”不能声称必然已崩溃；能力可用时inline二次确认“结束未完成研究/确认结束此研究”，说明保留草稿/附件、不会重跑、活跃build正常完成不会覆盖。结束后失败记录可沿既有确认cleanup。旧cap hide；pending只读查询/显式同次retry确认；cancel只停止等待。默认折叠/禁用/任务doc隔离；safe中文措辞不要暴露rawerror。

Root：研究Skill导出 async verifyPresentationResearchRecord(value)→parsedRecord+SHA checks，用于既有c.record及C status/ACK，await后原current guards保留；补purehelper最小校验tests。Agent指引仅read_running、用户可结束、继续研究需新ledger_id；不新增工具。actualPC→Agent原fixture验证orphan、restart、sameID不重跑、新ID可创建、结束后cleanup、记录SHA篡改拒、旧cap默认shape。统一full/四端types/static/diff/交叉review≤2轮/sourcecommit→Officebuild版本→docs。总体64%，真实专业任务0/20，不把合成测试当真实PowerPoint验收。

## 执行结果

已完成 A/B/C/Root 单元及交叉独立复审。实际失败 read/export 反例先失败后修复通过；完整相关2932/2932、最终70/70、四端类型/静态通过。源码d42d3213，构建d42d32138698通过，详见本批产品进度报告。总体64%，真实任务0/20，未部署。
