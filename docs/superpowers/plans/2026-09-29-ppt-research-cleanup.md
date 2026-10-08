# 研究归档清理与引用保护

依据原方案持久项目、资料管线、恢复/幂等和容量治理；当前明确工程缺口为ResearchStore128记录/64MiB无释放。基线1936ac66；高保证隔离+三单元独立实现/交叉审查；不删除真实用户资料，本次测试使用合成目录。

## 必须保持

研究Record.version1/id/sequence/draftDigest不可重编号。只由用户通过UI明确确认删除未被项目引用的已结束记录；Agent不注册删除tool；running记录拒绝。删除只清理ResearchStore原研究草稿/证据，不删除原附件、PPT、交付包及导出副本。当前/可恢复历史计划、所有compile/production冻结plan保护；读取引用历史损坏则拒绝清理，不能当作无引用。未完成交付上传若声明research.json可保守返回busy；已完成ZIP独立保留原研究副本，不强制清除。

## 固定公共契约

ResearchHistory/Summary增加version2分支：{version:2,documentId,projectId,revision,totalRecords,lastSequence,records}；totalRecords为现存数≤128，records仍最近min(32,totalRecords)，序号递增允许删除后的间隙且≤lastSequence。V1原所有字段/严格连续序号/revision关系完全不变；V2revision正safeint，lastSequence≥totalRecords，不能以删除重置序号。ResearchRecord仍V1。只首次成功删除将State迁移V2，普通现有begin/finish继续V1（直到清理）。

共享exports PresentationResearchDeleteReceipt + parsePresentationResearchDeleteReceipt，精确形状{version:1,documentId,projectId,ledgerId,sequence,draftDigest,deleteId,deletedAt,revision}，id≤128/digestSHA/真实ISO日期/positive safeints。保存最多4096墓碑/删除回执，原ID永不复用；达到墓碑容量拒新删除并保留已有数据。记录数配额按现存≤128，64MiB不扩大。Store V2校验所有现存与墓碑ID/sequence不重复、序号集合覆盖1..lastSequence；只删除finished时revision=lastSequence+finishedActive+2*tombstoneCount，墓碑receipt.rev应真实范围。checksum/原草稿SHA/软链接/普通文件/原子发布保障保留。ACK相同deleteId+ledgerId+digest幂等可重试，其他参数冲突拒；删除后read/begin同ledgerId返回record_deleted，finish拒，不能复活。取消提交前拒写，提交后丢ACK通过receipt只读恢复。

PC operations：research_capabilities optional includeCleanup:true（默认旧精确{version:1,available:true}）；扩展响应{version:1,available:true,cleanupAvailable:true,historyVersions:[1,2]}。
research_list/research_build optional historyVersion:2；没有协商时V1仍exact原shape；存储已V2而请求未协商返回upgrade_required，不能伪造连续序号。research_read/latest Record仍V1不变。
research_delete {documentId,projectId,ledgerId,deleteId,expectedDraftDigest,expectedRevision}→receipt。
research_delete_status {documentId,projectId,deleteId}→receipt/not_found。
安全错误record_deleted/record_running/record_protected/cleanup_quota_exceeded/aborted/revision_conflict/request_conflict/invalid_state，不吐原文。

## 单元

A remote_acquisition_pc：shared research.ts History/SummaryV2及receipt parser +research-store.ts migration/delete/status与全部Store/contract测试。Store API deleteRecord(doc,project,expectedRevision,deleteId,ledgerId,expectedDraftDigest,signal?)，deletedReceipt(doc,project,deleteId)。保护由PC项目锁内执行（Store不猜其它存储）。测试旧V1读取/新V2重开/容量释放128/保留序号/旧ID无复活/同ID ACK重试/坏档案/信号/配额/窗口外旧记录。

B addin_build_version：PC research service操作/协商/原引用保护，接入presentation-service.ts同现有 `${resolve(userDataPath)}\0${projectId}` 全局项目锁，覆盖清理与save/restore/compile/production冻结顺序。使用现有PresentationStore current.plan+revisions读planRevision，productionHistory和history所有冻结plans；缺已声明历史snapshot时fail closed；无计划研究项目可以清理。删除与绑定并发：绑定先赢则保护，删除先赢则后绑定拒archive缺失。护读损坏与跨doc/project也测试，禁止重复造独立锁而破坏现有串行边界。必要将已有项目锁逻辑抽为共同小函数；未完成带research.json上传busy。readyZIP清理后恢复不依赖researcharchive认证（实际确认）。操作status/receipt不删其它存储。无Agentdelete入口。

C remote_acquisition_office：Office research-controller.ts/research-card.tsx及tests用户确认/清理phase/旧能力隐藏/摘要V2/身份/CAS/当前doc守卫/丢ACK读取status恢复/重开。UI仅已结束item删除，running解释不可清理。可inline二次确认范围，deleteId固定绑定该record/digest，用显式重试或status读取，不自动重发delete。导出副本/附件/交付包仍保留必须明确。Card中文原文、窄面板操作禁用、Cancel只停止等待不承诺撤销。

Root：Office research Skill只读/build客户端协商V2/旧PC能力fallback（includeCleanup invalid_request映射upgrade后仅回退一次旧cap，旧cap响应也可合法接受），不新增deleteTool；协商成功才发送historyVersion:2，strict扩展cap解析与errors。共享纯Office helper readPresentationResearchCapabilities(requestOperationCallback) root写presentation-research-capabilities.ts供C用，返回{available:boolean,cleanupAvailable:boolean,historyVersion?:2}，callback(operation,body?)→Promise<unknown>，caller处理doc/abort/epoch。旧cap精确/扩展精确校验；扩展请求upgrade_required时旧请求fallback。actual PC跨层build/archive保护/delete/restart/capfallback/oldrequest升级/绑定并发/ACK丢失。readonly原研究证据Record和摘要SHA不变。

## 核对与门槛

各单元RED→GREEN；三端types/changedlint+format/diff、完整相关回归、交叉review≤2轮；源码统一commit→Officebuild版本核对→阶段/进度文档commit。无push/deploy。整体64%，研究治理闭环不等于专业语义或真实宿主验收完成。

## 实施结果

全部单元及交叉复审完成；源码 e4358bff，完整2860/222及最终补充83/9通过，三端类型/静态/Office构建通过。详见 docs/product/wiswork-ppt-agent-research-cleanup-progress-2026-09-29.md。整体仍64%，未部署，真实专业任务0/20。
