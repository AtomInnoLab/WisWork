# 研究证据到逐页主张复核与交付问题闭环

基线 de4e0e36 / 源码 d7f690a4。依据原方案 §4.3、阶段 B/F/G、§11.7，研究完成不等于事实支持/权威/时效完成。原研究已绑定计划，但冲突/遗漏只有独立摘要，不进入逐页问题账本；主张原文工具还缺完整专业限定和绑定研究上下文。本批把真实证据和这些缺口接入已有复核/交付路径，不新增自动认证或全局写锁。

## A：引擎逐页研究问题

Owned packages/pptx-engine/src/presentation-delivery-report.ts +相应test（不改claim-evidence/production/UI）。DeliveryIssue增加可选 research: {ledgerId,sequence,draftDigest,researchClaimId?,relatedClaimIds:string[],sourceIds:string[]}；sourceIds为原研究ID，不冒充plan sourceId。严格解析与重建seed一致，只有实际绑定研究的对应页/主张可生成。

将binding findings按claim分组为 research_unmapped_claim、research_conflict_partner_omitted、research_source_reference_unselected、research_source_unavailable 四类 issue；另外对实际mapped fact所有正向/反向冲突（包括双方均映射到计划）产生 research_claim_conflict，完整相关原ID在relatedClaimIds中。每claim每code最多1项，不丢相关多个source/conflict；稳定issue id含页/claim/code，不能跨页串处置。needs_human用于冲突与遗漏；unverifiable用于未映射与来源不可用。digest纳入研究描述符/对应原结论、冲突方、来源证据/引用上下文，不将supported review自动关闭研究问题。解释/暂缓沿用existingIssueLedger，仍不等于resolved。源变化或原记录内容变化应使旧处置stale/open。原无ref report shape行为保持；bound最多768 page issues（旧608+5*32），8MiB报告上限不变。Markdown保留原IDs/全文context，strict parser拒绝伪造research context。小RED→实现→targeted/types/lint，勿commit/全量。

## B：真实主张证据的研究与专业限定

Owned engine presentation-claim-evidence.ts +tests、PC presentation-production.ts +test、新pure helper可选。Evidence增optional documentId、claim（完整原Plan claim，source.asOf optional），research?:{binding: NonNullable<PresentationPlan['research']>,record:PresentationResearchRecord,findings:PresentationResearchBindingFinding[]}。只有绑定plan才返回这些新增完整字段（无binding旧shape不改）；读取冻结production.plan指定archive，document/project/descriptor与映射已由readBound校验，findings仅当前claim。禁止latest。严格Evidence parser检查record/binding schema/identity、claimId/statement/sourceRefs一致、full claim合法、选定source映射原研究原文与专业限定、相关finding references；客户端同步parser不宣称归档真实性，实际review写前重新构建证据及SHA。

完整研究Record保留全部冲突双方/未选来源/声明sourceTier与confidence/原retrievedAt，不从completed推断认证。canonical evidence digest覆盖所有新增上下文；因此旧不含研究的digest不能为新bound evidence存supported review。旧已保存review继续历史只读，不能伪造成新核验。仅bound Evidence总量512KiB（old256KiB保留），原文window仍8000 UTF16、附件/ZIP/报告上限不扩大。PC bounded response与tool路径相应容量由root负责。损坏archive写前安全失败、取消/late不能持久化review。小RED→真实PC对比原研究A与后续B/上下文篡改/restart→targeted/types/lint，不改report/UI/root Office production，不commit/全量。

## C：插件逐页研究问题界面

Owned Office src/agent/presentation-delivery-report-card.tsx +相关tests（不改engine/PC/root tools）。基于真实DeliveryIssue.research与report.research展示中文原因、原研究ID/版本、对应原结论、冲突方陈述/来源/限定、未选/不可用来源；默认折叠details，用户可说明/暂缓/重新打开，沿用同一issue.id/digest。页面不同的相同claim不会共享action。unsupported/needs_human不展示完成；已解释仍保留缺口。老报告不显示新section。小RED→实现→actual click/UI tests、types/lint，不commit/全量。

## Root：实际工具与跨层验证、统一交付

Office presentation-production.ts原claim evidence tool接受bound512KiB，检查documentId/claim/research actualdigest；原返回resource附件与明文摘要显示完整scope/专业限定/冲突/声明tier，并指示真正原文语义review，不按摘要/声明tier认定权威，修改稿后读最新准确窗口再record review。Prompt与工具description保持old兼容。

Actual cross：原附件→research A含冲突双方→mapped plan/frozen production→evidence包含A/限定/原文→research B/plan更新仍A→真实supported review→报告仍有冲突/遗漏issues→实际issue action解释/暂缓/reopen并restart恢复，supported不隐式resolve；异页同claim隔离，篡改evidence摘要/专业限定不写review，badarchive不写。旧无ref exact shapes、source/evidence provenance/locators/canonical digest与claimreview回归必须保持。

按TDD/ponytail/verification与subagent技能，将三项独立单元交叉审查，root统一完整relatedVitest、三端types、所有changedTS lint/format/diff；sourcecommit后Office生产build版本，再docs stage+整体台账。最多2review修正轮。无需上传/部署，不进行未授权实机验收。回滚源码不删研究/来源/已有review或issue action；新增字段optional旧记录可读。总体64%直到原退出门槛实际满足，不能用新增issue或绿测试冒充来源已核验。

## 实施结果

源码 99880bdb（13 文件、1437 行新增/33 删除）；完整相关2750/2750（210 文件）、最终定向53/53、三端types/changedTS lint/format/diff通过，交叉独立审查无遗留问题。共用 presentationResearchClaimBindingFindings 被全计划与evidence parser复用；Office持久缓存前重算研究草稿SHA和全证据摘要。真实cross含A→B/planB后freezeA、supported保留问题、逐页处置重开、伪造旧摘要拒绝及actual >256KiB完整原记录。Office build9.67秒/99880bdbefb7，未部署。总体64%不变，专业来源/时效及实机退出门槛仍未完成。
