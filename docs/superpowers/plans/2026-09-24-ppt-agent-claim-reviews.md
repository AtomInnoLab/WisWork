# 主张证据复核记录

基线774295a0。按原方案§4.3/O4推进；用户明确暂不实机验收，本批仅工程实现和自动化测试。沿用隔离分支、高保证/TDD/独立审查/全仓验证；整体估算42%，专业任务0/20，不修改原方案。

## 设计

持久化Agent对一个冻结page/claim/source证据窗口的判断，outcome为supported/contradicted/insufficient_evidence，notes非空<=2000；reviewer固定agent。这是审阅记录，不是自动事实通过，不改变原plan.reviewStatus、内容预检或宿主QA。每reviewId不可变、相同请求幂等、不同内容request_conflict。

共享新增presentation-claim-review.ts(export ./presentation-claim-review)：canonical evidence content helper presentationClaimEvidenceContent(evidence)（所有已验证证据字段稳定序列化，属性顺序不影响；明确含原文窗口/statement/source/摘要/检查结果），SHA256由PC Node/browser分别计算。PresentationClaimReview报告字段全部必需：{version:1,projectId,requestId,reviewId,planRevision,inputDigest,planDigest,pageId,claimId,sourceId,attachmentId,offset,maxChars,evidenceDigest,outcome,notes,reviewer:'agent',createdAt,checks:{support:'agent_reviewed',sourceAuthority:'not_verified',timeliness:'not_verified',host:'not_checked'}}。严格parser字段/ID/范围/日期/固定checks；正文notes为可展示文本，禁XML控制字符，不能执行。

PC操作production_record_claim_review：必需operation/documentId/projectId/requestId/pageId/claimId/sourceId/offset/maxChars/evidenceDigest/reviewId/outcome/notes。在同project串行锁中复用production_claim_evidence重新读取并计算digest；不匹配报evidence_changed；取消或变化不写。重新读取成功后原子持久保存。production_read_claim_review：operation/documentId/projectId/requestId/reviewId，返回历史记录，不重新声明证据仍有效，不需附件可用。

Store不引入pptx-engine依赖。新增saveClaimReview(projectId,documentId,requestId,reviewId,review:unknown)、claimReview(projectId,documentId,requestId,reviewId)。返回PresentationClaimReviewRecord envelope {version:1,projectId,documentId,requestId,reviewId,inputDigest,planDigest,planRevision,createdAt,review:unknown,reviewDigest}；review是{pageId,claimId,sourceId,attachmentId,offset,maxChars,evidenceDigest,outcome,notes,reviewer:'agent'}。基于存在production验证doc和frozen摘要。单production一个claim-reviews-HASH.json数组，最多32记录、256KiB；单review JSON<=8KiB，验证JSON无危险键、完整摘要、唯一reviewId/时间/frozen binding；相同内容返回旧记录，不更新createdAt；冲突request_conflict，容量quota_exceeded。复用原子write、read路径保护，不更改production记录。独立review文件不会进入production/receipt扫描。

插件复用production skill。read_presentation_claim_evidence成功后在内存保存最多16条最近证据的digest，key为完整doc/project/request/page/claim/source/offset/maxChars；clear清理。record_presentation_claim_review输入同evidence读入参+review_id/outcome/notes（digest不由模型提供），必须同会话先read才能记录，缺失返回presentation_evidence_read_required。调用PC传内部digest，严格校验回执与请求/绑定；记录完可同内容幂等重试，read历史不会授权写。read_presentation_claim_review只需project_id/request_id/review_id；历史读取不变更活动成果、VFS、QA、最近项目。写记录也是不变更宿主，不需要宿主确认弹窗。返回明确reviewer及未核验边界。旧PC升级提示。

## 分工与验证

store agent owns project-store/src/presentation-store.ts、index export、new tests。backend agent owns sharedcontract/package export、PC service/production+tests。root owns插件runtime路由/工具/tests、真实上传→读取证据→记录→重启读取integration、文档。各自限定commit，接口按上文。

TDD覆盖重复/冲突/容量、跨doc/任务、文件篡改/重启、证据digest变化不写、取消/失效读取不授权、伪造报告拒绝、历史读取不刷新信任。独立审查后全仓test/typecheck/lint/format/licenses，构建Addin/Shell。不做实机验收。

## 回滚

新增独立review文件，无旧数据迁移，旧插件忽略；回退保留记录。保留现有PC单进程project串行化边界，不声称跨进程CAS。对一份证据的review不代表整个claim所有来源均支持，更不代表真实Office内容/保存重开通过。
