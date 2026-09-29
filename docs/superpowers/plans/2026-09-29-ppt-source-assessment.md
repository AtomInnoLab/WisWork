# 来源权威性、时效与适用范围的可追溯 Agent 复核

基线 HEAD 1dd96d4a / source99880bdb。原方案 §4.3/§4.3.1、专业证据审阅、阶段F/G与Quality Pipeline。上一批专业限定/研究冲突已贯通，但每个来源的权威性/时效判断尚不可结构化持久化。本批复用不可变ClaimReview和原文窗口，不另建重复账本，不宣称认证、不创建全局写锁，所有全局checks仍not_verified。原方案退出门槛/20真实任务不变。

## Shared exact contract (B owns)

新增 browser-safe packages/project-store/src/presentation-source-assessment.ts +subpath export。PresentationSourceAssessment = {
 scope:string(1..400),
 authority:{outcome:'appropriate_for_claim'|'insufficient_authority'|'uncertain',sourceTier:'primary'|'authoritative_secondary'|'secondary'|'unverified',reason:string(1..600)},
 timeliness:{outcome:'current_for_claim'|'historical_only'|'superseded'|'uncertain',referenceDate:YYYY-MM-DD,claimAsOf?:string(1..100),sourceAsOf?:string(1..100),reason:string(1..600)},
 jurisdiction?:{claimJurisdiction:string(1..400),outcome:'applicable'|'mismatch'|'uncertain',reason:string(1..600)},
 basis:{offset:integer0..1_000_000,text:string(1..600)}[] max4
}。

exports PRESENTATION_SOURCE_ASSESSMENT_SCHEMA（普通JSON Schema兼容engineSchema与Agent inputSchema），parsePresentationSourceAssessment(value)，assertPresentationSourceAssessmentBasis(assessment,window:{offset:number,text:string})。严格keys/enums/date真实日历/UTF8总量<=16KiB/非空reason、重复basis拒绝；basis原文数据保留formfeed等原始code units，scope/reason/范围拒绝XML非法控制与孤立surrogate。basis absolute UTF16 offsets必须全在实际读窗口、text逐字相同。任一正向 authority appropriate/current timeliness/applicable jurisdiction要求至少1条basis；全uncertain可无basis以记录无可核验资料。literal存在只证明文字存在，不自动证明其内容正确或来源真实。

## A：引擎review与逐页来源问题

Owned packages/pptx-engine/src/presentation-claim-review.ts、presentation-delivery-report.ts +tests（不改Shared Store/PC/Office）。PresentationClaimReview sourceAssessment?:PresentationSourceAssessment；parser借shared schema/parse，旧无optional exact形状与check完全不变。保存review仍Agent，checks sourceAuthority/timeliness保持not_verified，不允许input自填通过。

有plan.research或任一report.review sourceAssessment时，逐claim/source读取所有对应历史assessments，不last-wins。每维至多1条结构化issue：source_authority_review_missing/uncertain/insufficient/mixed（authority outcome/tier有差异即mixed）；source_timeliness_review_missing/uncertain/historical_only/superseded/mixed（outcomes不同或referenceDate/claimAsOf/sourceAsOf范围不同即mixed，明确是判断/比较框架不同非自动事实矛盾）；claim.jurisdiction存在则source_jurisdiction_review_missing/uncertain/mismatch/mixed。正向一致判断不加该维问题，但全局checks依旧notverified，页/claim type不变。已有unbound无assessment report全shape/issue/digest兼容。

问题种子沿用page/claim/source ID，digest覆盖本scope所有带assessment的review及原claim/source/frozen研究引用，不让旧支持窗口、说明或暂缓关闭新来源问题。引用或判断变化旧action stale/open；跨页/任务隔离。bound/assessed页最多1056 =旧608+5*32+3*3*32，原unbound未assessed608，report8MiB不变。Markdown清楚保存assessment全文/原basis/refdate及历史判断范围，不能从sourceTier/retrievalAt/asOf标签推断认证。RED→实现→targeted/types/lint；不commit/full/委派。

## B：shared +真实存储与PC复核写入

Owned newshared contract/export/tests、project-store/presentation-store.ts ClaimReview optionalpersist及tests、apps/shell/src/main/presentation-service.ts+presentation-production.ts+PCtests（不改engine review/report/UI/Office tools）。record_presentation_claim_review新增optional sourceAssessment request camelCase；旧request strictshape继续有效。service入参验证新schema，PC重建actual evidence+sameSHA后断言basis实际窗口，timeliness.claimAsOf/sourceAsOf必须与frozen真实plan claim/source标签严格一致（未提供标签也不能补猜）；有jurisdiction assessment其claimJurisdiction必须等于该frozenclaim.jurisdiction，不存在不可编造。检查在write前，abort/clear/丢ACKsameID幂等，sameID不同assessment拒绝。

Store reviewDigest纳入optional完整assessment，old review原8KiB/newassessed24KiB per-record，32records/total256KiB不扩大；read/restart/checksum/corrupt异常严格，raw未知错误不泄露。无新增外部访问/automaticcertificates。

## C：实际来源复核界面

Owned Office delivery-report-card.tsx +tests（不改engine/PC/Root tools）。增加所有13来源issue code中文解释（authority4/time5/jurisdiction4），显示该页/claim/source所有相关带assessment的历史判断（原reviewID/window/createdAt/声明tier/判断reason/referenceDate/asOf/辖域/basis），不择最新和不把不同框架直接称事实矛盾。默认折叠；既有open/deferred/explained action按钮复用IDs/digests/CAS，隐藏noassessment历史额外detail，old无ref/assessment报告保持。正向统一仍标历史Agent判断、不显示全局通过。实际click/跨页/任务/old覆盖；RED→实施→targeted/types/lint，不commit/full/委派。

## Root：工具schema与实际跨层

Office production.ts record_presentation_claim_review optional snake source_assessment，经sharedschema+parser+basis缓存窗口前校验，body sourceAssessment。严格full返回与所提交assessment canonical相等（旧PC忽略不得误称写成功），read historical parse保存全assessment。Prompt明确原文/层级/时点/适用范围的语义判断是Agent意见、compareframe需copy真实frozenlabel、不用found/sourceTier/retrievedAt当认证；旧PC不支持optional返回upgrade安全提示，保持old无assessment工具。报告/当前宿主ZIP由既有完整reviews自然附该immutable data。

实际cross：真实附件/researchA/plan/frozen/evidence→首条合适+current/applicable assessment有真实basis→后续uncertain/历史/其他tier/referenceDate形成mixed→两条都保留→报告/处置/重开，支持与正向判断仍不关闭研究冲突或project.completed。伪造offset/text/时点/辖域/未知字段/duplicatebasis写前拒绝；lostACKsameidreadonly找回、同id改变拒绝、oldcompat；当前host包可附source assessments且PC原report验证不降级。先RED，root完整relatedVitest/三端types/changedTS静态+独立交叉review<=2rounds，sourcecommit后Office build版本，再docs+总体台账。回滚源码不删原件/旧review/issuehistory；没有上传/部署或实机PowerPoint。总体64%直到原模块exit实证满足。

## 实施结果

全部四个工程单元完成，独立交叉审查通过；完整相关2780/2780、三端类型与静态检查通过，源码e5fcc8d3，插件构建e5fcc8d371a4。阶段报告记录实际范围和剩余退出门槛，总体64%。
