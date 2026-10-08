# 科研、法律、金融专业上下文贯通

依据原方案 §4.3.1，不改变既有验收状态。基线 1bc7b7f8，隔离分支 codex/ppt-agent-implementation；跨公开契约采用已有高保证独立实现与交叉审查。

## 契约

ResearchFact 与 PlanClaim 可选 professionalContext，使用一个 browser-safe shared schema/parser。对象为 domain:'science'|'law'|'finance' 加该领域可选字段；仅 domain 必填，缺字段成为结构化警告而不是拒绝材料录入。所有文本1..800 XML兼容，日期 YYYY-MM-DD 严格真实日历，整体≤16KiB。

- science: materialKind:'paper'|'dataset'|'standard'|'institution'；publicationId（DOI/官方数据集/标准标识）、version、sample、method、statisticalBasis、limitations。
- law: materialKind:'statute'|'case'|'regulation'|'contract'；jurisdiction、effectLevel、effectiveFrom、effectiveUntil、applicabilityDate、caseNumber、originalLocation、limitations。
- finance: materialKind:'disclosure'|'financial_statement'|'ir'|'market_data'；reportingPeriod、asOf、currency、unit、accountingBasis、formula、limitations。

schema展示所有字段；runtime按domain严格拒绝跨领域字段/错误kind/undefined/未知字段。部分上下文允许缺失但不得填猜。两日期若均有且 effectiveUntil<effectiveFrom 拒绝。导出 PROFESSIONAL_CONTEXT_SCHEMA, PresentationProfessionalContext, parsePresentationProfessionalContext, presentationProfessionalContextMissingFields。

必要字段：science materialKind/publicationId/version/sample/method/statisticalBasis/limitations；law materialKind/jurisdiction/effectLevel/applicabilityDate/originalLocation/limitations，非contract还需effectiveFrom，case另需caseNumber；finance materialKind/reportingPeriod/asOf/currency/unit/accountingBasis/limitations，calculation类需formula（helper可选claimType参数）。case/contract无effectiveUntil不自动推断仍有效。

冻结映射与主张证据 professionalContext 必须逐字段canonical等值于指定原事实；去掉上下文不能保存/冻结/写复核。旧没有字段的形状和摘要保持。通用 jurisdiction/asOf/calculation 不得隐式覆盖专业原值，矛盾可录入但产生人工警告。

## 单元

A Shared +Research：新 professional-context.ts/export/契约测试，research.ts fact/schema/parser与归档测试；Office research工具说明/MD完整上下文。已有归档SHA和限额不变。

B Engine：Planclaim schema/type/parseroptional字段，research-binding 与 claim-evidence 精确等值/解析，deliveryreport专业警告及MD，测试。代码 professional_context_incomplete（单claim聚合missing）、professional_source_secondary（绑定原fact tier secondary/unverified）、professional_legal_rule_inactive（明确applicabilityDate<effectiveFrom或>effectiveUntil；日期边界inclusive）、professional_jurisdiction_mismatch（通用与专业均有且逐字不同）、professional_financial_time_mixed（通用asOf与专业asOf均有且不同）、professional_financial_unit_mismatch、professional_financial_currency_mismatch（仅存在calculation相应值且不同）。全部needs_human或缺上下文unverifiable；每claim/code≤1，全部不认证。现绑定/assessed1056+7*32=1280，当有professionalContext启用1280，否则旧608/1056保持；report8MiB保持。issue digest覆盖完整professionalContext，原完整claim已有覆盖应验证。旧unbound evidence返回shape保持，但有professionalContext时必须携带fullclaim/doc，并进入digest。

C UI：ReportCard专业原因中文、默认折叠逐页完整专业上下文及missingfields，正向完整无issue也可见；research原fact展示context；old隐藏。真实click/action隔离。

Root：实际PC production evidence对有professionalContext无research任务也发送fullclaim/doc；Office response窗口完整digest/专业指引、实际跨层三领域与原research映射篡改/恢复/交付包测试。按相同协议不可缺omit fields伪造成功；不要重开重算latest代替。

## 验证和结尾

失败先行各单元定向；交叉独立审查最多两轮；root完整相关回归/三端types/改动lint+format+diffcheck；源码统一commit→Officebuild版本验证→阶段/台账docscommit。无push/deploy。总体64%，真实专业PowerPoint验收未替代。

## 实施结果

四单元完成并通过交叉独立审查；full2809/2809、三端类型/静态通过，源码cf1463a6。Root实际跨层13/13，明确旧shape无法仅凭响应证明完整性，但实际PC写前SHA守卫拒绝删除上下文的新判断。当前仍64%，本批非专业语义/真实宿主验收完成。
