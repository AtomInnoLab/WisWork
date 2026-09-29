# ResearchLedger 到制作计划的精确版本绑定

依据原方案 §4.3、阶段 B/C/F/G 与可追溯主张要求。当前独立研究已先于计划可保存，但计划与研究之间缺机器契约，交付只能附最近历史。本批让实际保存计划/冻结生产/报告/当前宿主 ZIP 引用同一不可变研究记录，并显式保留引用缺口，不提升为事实认证。沿用已授权实现分支、原研究归档与现有计划/生产/报告路径，不新增外部调用或新的全局门禁。

## 契约与引擎单元 A

- `PresentationPlan` optional `research: {ledgerId:string,sequence:positive integer,draftDigest:sha256,sources:{sourceId:plan source ID,researchSourceId:research source ID}[],claims:{claimId:plan claim ID,researchClaimId:research claim ID}[]}`。strict keys、映射无重复plan ID、映射plan目标存在；研究ID允许128，计划ID保持原限制。不要求研究来源ID等于计划来源ID，显式映射避免截断/重命名丢失身份。
- 新 engine纯模块 `presentation-research-binding.ts` 及export subpath，类型 `PresentationResearchBindingFinding`，functions `assertPresentationResearchBinding(plan,record)` 与 `presentationResearchBindingFindings(plan,record)`。
- assert：Record strict parser+completed，projectId、ledgerId、sequence、draftDigest一致；映射source/claim均实际存在。引用原文 URI、snapshotAttachmentId、excerpt、locator、asOf完全相等（标题和ID可改名）；mapped claim statement/type/asOf/jurisdiction/calculation基础formula/inputs/unit/currency完全相等，计算reproduction可后续补充；每mapped claim的plan sourceIds都映射到该research fact源集合。研究可多于计划3引用上限，未选引用显式finding而非静默丢失。confidence允许声明变动，但仍needs_review且不推断认证。
- Findings deterministic bounded：`{code:'unmapped_claim'|'omitted_conflict_partner'|'unselected_source_ref'|'source_unavailable',claimId,researchClaimId?,relatedResearchClaimId?,sourceId?}`。未映射计划主张、选了一方而对方没有映射、研究引用没有选用、对应原文未ready/found，都保留。它们是计划引用范围的需复核项，不是新增全局写锁。保留双方完整Record供审阅；不自动挑叙事便利一方。
- optional frozen DeliveryReport `research?:{record:PresentationResearchRecord,findings:PresentationResearchBindingFinding[]}`；当plan.research存在必需精确Record，未绑定计划不应冒充绑定报告。parse报告核对doc/project/实际冻结plan、descriptor和重算findings，拒绝错误Record/伪造findings/仅摘要。build input新增researchRecord?:Record，Markdown输出明确冻结研究ID、双方冲突与缺口，checks原未核验值不变。报告内容预算继续有界，不吞missing原归档。
- owned files A：engine presentation-plan.ts、新binding helper、presentation-delivery-report.ts、package export及相关tests。不要改PC/Office/sharedStore。

## PC 单元 B

- actual save_plan 在已有project锁下读取原独立ResearchStore archive，检查actual document/project、Record completed、实际draftSHA（Store已验证），复用engineassert，不把最新结果代替指定记录；失败不保存计划。不要求绑定最新research，支持明确旧研究版本；旧计划无field零额外要求。
- restore_plan/save历史恢复、production_begin/frozen生产输入新任务读取相同原研究核对；基于旧绑定恢复仍用旧record，新研究不清空旧计划/资产/完成页。已有冻结任务不用最新记录替代，损坏/缺失归档明确失败，禁止伪造或自动重新研究。
- handlePresentationDeliveryReport读冻结production.plan.plan.research，并callback读取指定原Record交给enginebuild；未绑定仍保持旧report完全shape兼容。service传可选`readResearch(ledgerId):Promise<Record>` callback，直接独立handler测试旧签名可继续无ref。safe errors `research_binding_invalid`/`research_unavailable`，unknown/raw错误不泄露；保存回执丢失相同plan CAS/幂等仍返回原版本。
- owned B：PC service、新 `presentation-research-plan-binding.ts` 可小helper、PC report handler与tests。不要修改engine/Office/rootbundle。

## Office 单元 C

- 既有save_presentation_plan schema自动含research，systemPrompt解释从read_research_ledger读取实际record后显式映射计划source/claim，保持原摘录与专业限定，不猜ID/digest，不从completed推断核验。
- ProjectController/Card按当前计划展示绑定研究序号、映射数量、范围说明，ID/摘要可折叠；提供可选 `readBoundResearch?()` readonly action，options callback `readResearchRecord?(projectId,ledgerId,signal?:AbortSignal):Promise<void>`，Runtime调用独立ResearchController选择明确project并读明确ledgerId，不能用latest代替。取消/文档变化/旧PC安全守卫，与现有研究workspace结合。根tools已有read_research_ledger，复用不新增读工具；已有ResearchCard展示真实目标Record。
- 实际报告Card显示plan-bound历史研究与finding中文说明/原ID，冲突/缺口不得自动closed或QApassed；旧未绑定报告隐藏该section。保持原计划调整、撤销与恢复，不自动替换人工选定研究。
- owned C：Office planning prompt、ProjectController/Card、HostRuntime callback、DeliveryReportCard与tests。不得编辑engine/PC/rootbundle。

## 根单元 D / 联合验证

- 当前宿主Bundle优先使用 report.research.record；只有无binding才读取最近历史research callback。显式绑定存在却缺record拒绝，不悄悄降级到最新。README解释计划精确绑定与仅项目历史的区别；PC bundle finish当冻结report.plan有binding时要求research.json与report.research.record完全相同且原archive匹配，防错误附录替换。
- 实际crosslayers：先保存研究A→创建含mapping计划→冻结production→新研究B→改变计划版本→读取旧production报告/宿主ZIP仍使用A；无计划依赖研究、保存/retry/重开、原文/claim/计算限定变化拒绝写、引用3子集/遗漏冲突findings保留、恢复旧计划精确绑定、坏归档/取消/旧无ref兼容、latest不得替换。
- 先RED then实现，各单元targeted/types/static，不自行commit/全量；根统一sourcecommit、完整相关回归/三端types/独立review<=2修正轮次、Office构建版本后docs提交。现有源文件引用不存在时feature不可假装完成；旧缓存/计划/production默认无ref保持形状与行为。撤销源码保留archive，无外部上传/部署/实机授权变化；原方案20真实任务/全部退出门槛及成熟度口径不变。

## 实施与验证结论

源码 `d7f690a4`（26 文件），完整相关 2735/2735、最终定向 7/7、三端类型/改动源码静态检查通过，交叉独立复审无遗留重要问题；Office 构建 10.14 秒、buildId d7f690a486da，未部署。实现中补齐 legacy resume 的冻结研究校验，并在异步归档读取后关闭保存/冻结的取消窗口。Office callback 采用可选 AbortSignal 实际贯通取消。阶段详情与仍未完成的专业/实机门槛见产品进度记录；总体 64% 不变。
