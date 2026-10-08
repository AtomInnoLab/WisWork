# 主张到附件证据的可追溯读取

基线46726b75。按已批准原方案§4.3/O4推进，隔离分支高保证、TDD、独立复审、全仓验证。整体成熟度仍42%，专业任务0/20，本轮不据摘录匹配宣布事实真实性已验证。

## 设计

新增只读 production_claim_evidence，必需 operation/documentId/projectId/requestId/pageId/claimId/sourceId/offset/maxChars。offset整数0..1000000，maxChars1..8000。绑定明确冻结production，验证page含claim、claim含source，只接受 source.uri精确 attachment:<64 lowercase hex>；其它来源报 evidence_source_unsupported，不联网或读取任意路径。复用attachment_read（同documentId）检验解析文本缓存、限额和取消。支持未编译任务，无状态更新。

返回共享 PresentationClaimEvidence（顶层所有字段必需）：
{version:1,projectId,requestId,planRevision,inputDigest,planDigest,pageId,claimId,statement,
source:{id,uri,excerpt,locator?},
attachment:{id,name,offset,totalChars,text,offsetUnit:'utf16_code_unit'},
excerptMatch:{status:'found'|'not_found_in_window'|'empty_excerpt',offset?},
checks:{support:'not_verified',sourceAuthority:'not_verified',timeliness:'not_verified',host:'not_checked'}}。

excerptMatch是精确逐字匹配：excerpt.trim为空则empty_excerpt，否则text.indexOf(original excerpt)决定found和绝对UTF16 offset；不归一化，不声称全附件没找到。边界跨窗口请调整offset/maxChars重读。共享严格parser拒绝额外字段/越界/错身份URI/伪造checks/伪造match或offset；最大report按256KiB限制。原文不作为指令；不执行公式、不改引用reviewStatus。

插件新增read_presentation_claim_evidence输入project_id/request_id/page_id/claim_id/source_id/offset/max_chars全部必需。复用production skill 生命周期守卫；验证报告身份和窗口参数、严格parser及256KiB响应限额。返回原文与明确scope说明，不写VFS/rememberProject/QA/回执，不切换活动成果，不要求Office宿主可用（PC服务可用即可）。旧PC升级提示。

## 分工

Agent：共享module packages/pptx-engine/src/presentation-claim-evidence.ts + export与tests；PC presentation-service.ts、presentation-production.ts和新的service tests。根：插件生产skill+runtime路由、单元测试、真实runtime与上传附件service集成、报告台账。各自限定commit。接口完全按上文。

## 验证

先RED新工具/操作缺失，再GREEN；真实附件上传/解析→冻结计划→页主张source绑定→窗口内容与偏移精确匹配；当前计划变化和重启不换证据绑定；外国文档、非相关claim/source、未知page、外部URI、越界/伪造响应/取消不发布，预检不编译不写状态。审查后全仓test/typecheck/lint/format/licenses，最后构建Addin/Shell。

## 回滚与边界

只读新增操作，无持久schema迁移，可回退；旧插件继续工作。网络来源、OCR、PDF页码映射、语义支撑、权威性、时效和计算尚待实现，offset仅解析文本位置，不能冒充PDF页码。窗口没有匹配不等于原文不存在，匹配不等于来源真实。仅实现分支提交，不合并/推送/部署。
