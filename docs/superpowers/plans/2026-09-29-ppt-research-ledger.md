# 制作计划前的独立 ResearchLedger

依据原方案 §4.3、阶段 B、§7.1 和 build_research_ledger 工具。研究账本必须先于计划和生产任务，不能把现有逐页生产审查改名为研究。沿用已授权隔离实现分支与现有附件管线，不新建外部检索服务。Agent 通过既有搜索及资料导入读取原始材料后提交结构化研究结论；PC 保存并复核实际文档绑定的原文摘录，不凭搜索摘要认定事实通过。事实/引文/计算/判断/假设及冲突均保留，权威性与时效性尚未通过时明确标注。

## 固定契约

共享纯模块 `packages/project-store/src/presentation-research.ts` 子路径 `@wiswork/project-store/presentation-research`：

- `PresentationResearchDraft`：`{scope:string,sources:Source[],facts:Fact[]}`。
- Source：`{id,title,uri,snapshotAttachmentId?,excerpt,locator?,asOf?}`；与现有计划来源字段/限制一致。uri attachment:SHA 或公开HTTP(S)URI，网页必须 snapshotAttachmentId 才能匹配原文，未提供记录为缺失而非编造。
- Fact：`{claimId,statement,type:'fact'|'quote'|'calculation'|'judgment'|'assumption',sourceRefs:string[],sourceTier:'primary'|'authoritative_secondary'|'secondary'|'unverified',slideRefs:string[],confidence:'high'|'medium'|'low',reviewStatus:'needs_review',conflictsWith:string[],asOf?,jurisdiction?,calculation?:{formula,inputs:string[],unit?,currency?}}`。使用原 ClaimRecord 名称；slideRefs 是建议使用页面，未声称实际宿主存在。sourceRefs/conflictsWith 引用必须存在，不重复/自引用，冲突双方不可丢弃。事实无来源允许但为未核验/缺口，判断/假设显式区分；不能从声明的 sourceTier/confidence 推断认证。source/fact 最多各64，UTF8 draft<=256KiB，strict keys、字符串/日期/ID/URI/计算字段等。
- `PresentationResearchRecord`：`{version:1,documentId,projectId,id,sequence,draftDigest,draft,state:'running'|'completed'|'failed',startedAt,finishedAt?,sources?:Evidence[],error?:'aborted'|'source_unavailable'|'invalid_state',checks:{scope:'research_draft',support:'not_verified',sourceAuthority:'not_verified',timeliness:'not_verified'}}`。
- Evidence：`{sourceId,attachmentId?,status:'found'|'not_found'|'empty_excerpt'|'not_ready'|'unsupported'|'missing'|'source_mismatch',offset?,locator?,provenance:'user_supplied'|'fetched_url_matched'|'unavailable',retrievedAt?:ISO string,sha256?,parsedTextSha256?}`。真实元数据决定 provenance，检索时间只来自网页原快照，用户材料不给虚构检索时间；开始/结束是研究整理操作时间。
- `PresentationResearchHistory`：`{version:1,documentId,projectId,revision,totalRecords,records:PresentationResearchRecord[]}`，最近32记录，全球sequence连续窗口；完整记录按id保留供只读恢复，不默默覆写。每record<=512KiB，公开history<=512KiB时可摘要？固定这里 history 为全record数组因此history需 bounded max16MiB，浏览器历史list不可直接history；另摘要HistorySummary为`{version,documentId,projectId,revision,totalRecords,records:[{id,sequence,draftDigest,state,startedAt,finishedAt?,sourceCount,factCount,conflictCount,error?}]}`<=64KiB。parser names `parsePresentationResearchDraft/Record/History/Summary`。

## PC 单元 A

新增独立 store `presentation-research-store.ts`，子路径Node-only export。用户本机 `presentation-research/{hash(doc)}/{hash(project)}`，原子checksum文件、符号链接拒绝、有界读取、严格临时清理及每文档项目序列化；账本最近32，完整records持久保存按id可读，累积最多128条/project、64MiB显式quota（达到限额拒绝新build，既有可读，不自动删除）。Global revision CAS，begin(expectedRevision,id,draft)第一次创建running并递增revision/total，相同id+draft幂等返回原记录（旧expectedRevision重试允许）；不同draft同id拒绝。finish更新原record与history，completion/failed不伪造事实通过；并发 CAS严格，丢回执只读恢复。failed相同id重试返回原失败，要新id显式新尝试；running重开只读保持未决，不自动继续或重放。

PC `presentation-research.ts` handler，通过既有presentation.v1早期独立route：

- `research_build {documentId,projectId,ledgerId,expectedRevision,draft}` -> `{history:Summary,record:Record}`；不要求保存计划/production。begin先persist，再复用现有 auditPresentationSources 或共享原文匹配实际attachment服务，核对元数据ID/sha/textdigest/sourceURLhash；支持无快照/缺失/不匹配等结构化缺口，捕获安全失败，no raw errors。对用户上传 originals标user_supplied；完整网页来源hash匹配才fetched_url_matched，未准备不宣称ready证据。
- `research_list {documentId,projectId}` -> Summary；可并发读到实际running start，无自动重放。
- `research_read {documentId,projectId,ledgerId}` -> Record。
- `research_latest {documentId,projectId}` -> `{record:Record|null}`，只读返回完整archive最近completed，不能因32摘要窗口把更早有效结论丢弃。
- `research_delete {documentId,projectId,ledgerId}` -> `{ledgerId,deleted:true}`仅显式清理已completed/failed archived，running禁止，保留窗口中的记录不能删（避免窗口序号篡改）；若实现复杂，本批先不启用delete，128/64MiB上限明确不可新写但可旧读，后续明确用户归档清理。固定前3操作即可，不假装有UIcleanup。
- `research_capabilities {documentId}` -> `{version:1,available:true}`，独立于项目存在，可旧PC隐藏。状态optional `researchAvailable:true` planned/compiled供工作台恢复。列表绑定真实doc/project，其空间必须严格隔离；不访问原输入、不要求生产。不生成批准/事实验证/项目完成。

## Office 工具根单元 B

新 `createPresentationResearchSkill({available,request,documentId,vfs,lastProject?,rememberProject?,onChanged?})`→AgentSkill&{clear()}; tools `build_research_ledger` input `project_id,ledger_id,expected_revision,draft`、`read_research_ledger` input project_id,ledger_id、`list_research_ledgers` project_id、`export_research_ledger` project_id,ledger_id。strictparser前后校验，doc/epoch/signal/busy，真实PC read/list；旧PCupgrade提示。build返回完整record+summary，failed明确操作未完成，原事实/冲突保留；不要用state.completed宣称核验通过。记住project仅成功记录后，不能替代当前plan项目status存在。export原始JSON+humanMarkdown（scope、结论type、declaredtier/confidence、原文/可点source、每源缺口、conflict双方、needs_review和未验证项），VFS atomic。root ACP研究阶段映射与实际跨层附件资料→独立账本→双端重开测试。交付包在导出时读取明确匹配项目最近completed研究Record作为历史研究附录（本批可添加 optional `research.json/.md` 到共享ZIP允许字段，严格SHA文件与冻结evidence区别；不伪装与生产任务冻结绑定）。若无record不虚构，manifest checks不变，README明确research为项目历史整理结果、非当前宿主事实验证。

## Runtime/UI 单元 C

实际Runtime注册新ResearchSkill/工具dispatch/clear和可选callback读取。研究必须在无plan/production时可见：新增独立ResearchController（capabilities probe、lastProject或明确selectProject、真实list/read/export、doc/epoch/cancelguard），Runtime/App/WorkspaceUi接通独立Card，不只挂在依赖project status的旧控制器。Skill onChanged?(projectId):void 通知独立controller刷新。既有project controller可按需增加`research?:Summary`、`researchRecord?:Record`、可选`readResearchLedgers?()/readResearchLedger?(ledgerId)/exportResearchLedger?(ledgerId)`，按project researchAvailable才自动list；新无plan项目由Agent工具开始，list/export不依赖production。独立Card `presentation-research-card.tsx` 嵌入实际工作台，独立于ProjectCard status存在，默认折叠研究历史/事实详情，显示作用域、来源/结论数、冲突数、安全失败、running无终结不推断后台正在运行；每fact区分推断与事实、声明可信度不视作认证、来源链接/缺口。读取历史记录/导出入口，取消/doc/taskguard，失败不清空plan/assets，不用自动批准。旧PC隐藏。不改PC/shared/rootskill。

## 验证/提交

各单元失败先行测试+定向类型静态检查，不自行commit，根合并一次sourcecommit，相关全回归和三端types、独立审查最多两轮修复，生产构建后记录sourcebuildID。涵盖真实网页/用户原件身份、来源缺失/摘录/locator/来源不符、conflict双方、重复/CAS/重开/响应丢失、并发running只读、取消、quota/symlink/坏checksum、先于plan可写、旧PC、UI与ZIP研究附录。不新增外部上传/部署，测试合成原件/网页。独立缓存回退保留，原方案20真实任务和全部退出门槛不变；总体百分比仍按九模块成熟度证据更新。
