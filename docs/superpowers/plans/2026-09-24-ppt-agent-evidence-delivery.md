# 内容证据交付包：复算、处置、摘要与导出

基线02489123，既有隔离分支。用户要求按原方案一次推进多项关联能力，实机验收继续暂缓。对应原方案§4.3 Claim Ledger、阶段F/G、O4。高保证：TDD、子代理分工、独立审查、全仓验证；不改原方案、不引入外部依赖、不执行用户脚本、不联网。

## 交付与边界

1. 冻结计划可声明计算的机器输入；受限四则表达式复算，与声明值比较。结论仅为算术复现，不证明输入来源、单位换算、语义结论或时效。
2. 内容问题有持久处置历史：open/deferred/explained，需理由；explained仅记录说明，不关闭机器发现、不把Agent标为人审。重新打开问题保留历史。同一请求的发现摘要变了，旧处置显示stale且当前按open处理。
3. 逐页交付摘要整合既有内容预检、所有来源复核历史、复算和处置；已核验（仅算术复现）/待人工判断/无法核验分类明确。所有页面、claim/source关联完整；不改宿主QA。
4. Taskpane可查看选择任务的摘要、记录处置、导出JSON+Markdown到现有会话附件。JSON含冻结Plan/Claim Ledger/来源与复核记录/全部处置历史；Markdown含人可读证据与警告；链接与文本安全转义，不自动抓取URI。沿用既有PPTX下载/导入，不冒称Office验收。

## 契约（协作基准）

### 计算

PresentationPlan.claims[].calculation新增可选reproduction:{bindings:{name:string,inputIndex:number,value:number,sourceId:string}[],expected:number}。bindings数必须等于inputs，inputIndex唯一覆盖0..inputs.length-1，name唯一ASCII字母开头32字符以内（不允许prototype/constructor/**proto**），sourceId必须在该claim.sourceIds中，所有数值有限且abs<=1e12。旧plan仍可解析，不配置reproduction时保留not_configured。公式仅数值字面量（含科学计数）、绑定名、+ - * /、一元正负号和括号；不支持函数/属性/赋值/幂/脚本。最多256 tokens、32层括号、128运算；全部binding必须参与表达式；未知名/额外字符unsupported_expression；除零/非有限/abs>1e12 invalid_arithmetic。固定容差8*Number.EPSILON*max(1,abs(actual),abs(expected))，不接受任意用户放宽。结果{claimId,status:'not_configured'|'reproduced'|'mismatch'|'unsupported_expression'|'invalid_arithmetic',actual?,expected?,tolerance?,scope:'arithmetic_only'}。只支持非金融定点承诺的IEEE double近似算术；保留输入、公式、声明值与误差供追溯。

### 处置持久化（browser-safe @wiswork/project-store/presentation-issue）

IssueActionInput={actionId,issueId,issueDigest,state:'open'|'deferred'|'explained',note}，ID<=128 ASCII安全，digest64hex，note trim非空<=2000且XML安全字符。Record={...input,sequence,createdAt}；Ledger={version:1,projectId,documentId,requestId,inputDigest,planDigest,revision,actions:Record[]}。revision=actions.length=最后sequence；最多128条不截断，不删除旧动作；相同actionId/内容幂等，异内容request_conflict；同内容幂等先于expectedRevision检查；其它CAS冲突revision_conflict。Store issueActions(projectId,doc,req):Ledger（未存返回revision0空），appendIssueAction(projectId,doc,req,expectedRevision,input):Ledger。存储原子带摘要、验证生产请求身份/冻结digest、损坏拒绝。Store不负责证明issue存在，由PC重新生成报告并核对issueId+digest后写入；明确missing/stale→issue_changed。无旧文件迁移。

### 汇总（browser-safe @wiswork/pptx-engine/presentation-delivery-report）

parsePresentationDeliveryReport、buildPresentationDeliveryReport（可async）、presentationDeliveryMarkdown；具体类型由engine在实现开始即发布给协作者。固定根字段version1,projectId,documentId,requestId,planRevision,inputDigest,planDigest,plan, reviews:PresentationClaimReview[], issueLedger:Ledger, pages, checks。page:{pageId,title,productionState,calculations:CalculationResult[],issues:DeliveryIssue[]}。issue:{id,code,claimId,sourceId?,digest,category:'needs_human'|'unverifiable',disposition:{state,stale,actionId?}}；默认open/stalefalse，若该issue最后action digest变化→open/staletrue并保留actionId；历史未匹配当前issue的动作保留ledger不丢弃。id可使用冻结page/claim/source索引与code构成稳定ASCII短ID。digest=SHA256 canonical({inputDigest,planDigest,pageId,code,claimId,sourceId?,relevantReviews/calculation})，同source新增review改变相关问题digest，不使其它页/来源处置失效。

issue code包括既有content finding（复算reproduced时去掉calculation_not_reproduced；其它按结果code更精确）、claim_no_sources、source_review_missing/mixed/contradicted/insufficient、calculation_mismatch/unsupported/invalid_arithmetic；supported仍历史Agent判断，不等于来源真实。对于不可核算/无来源/缺少摘录等归unverifiable，其余needs_human。checks固定scope:'frozen_production',content:'needs_review',sourceAuthority:'not_verified',timeliness:'not_verified',host:'not_checked',roundTrip:'not_run'。全报告<=8MiB，每页issues<=512（按最坏组合核算），32pages，全plan192KiB仍有效。parser严格字段、唯一性、身份与claim/source/page归属、状态组合、复算结果（可根据plan重算）、处置和ledger一致性，不接受伪造机器passed。export禁止隐去未处置项或stale。

PC新增production_delivery_report精确{operation,documentId,projectId,requestId}，返回report；production_record_issue_action再加expectedRevision,action（上述input），在项目锁内重建报告/验证digest后append，返回更新report。取消在写前生效，写入后取消不声称未保存；支持幂等重试。复用冻结plan/deck和listClaimReviews，无自动重新读取附件，不继承父任务处置。

## 分工与文件

A engine：presentation-plan.ts及tests（可选reproduction校验）、新presentation-calculation.ts及tests、新presentation-delivery-report.ts及tests、package exports。不改既有content-check契约（汇总层消费/调整其finding）。共享reportbuilder输入见协作者约定，需要单一schema和明确checksum序列化；先发布类型，然后实现，TDD恶意表达式、除零/溢出、精度、旧计划、映射来源、原文/Markdown转义、报告篡改/处置stale/多来源/完整覆盖。限定提交。
B store/PC：新project-store/presentation-issue.ts+导出和store方法/tests；shell新presentation-delivery-report.ts+service路由/tests。和engine同步builder签名；复用现有ClaimReview构造，project锁内计算、write前signal检查；身份/变化冲突、CAS/幂等、重启/损坏/no-host写回归。限定提交。
C frontend：新skills/powerpoint/presentation-evidence-delivery.ts，projectcontroller/card、相关测试。工具read_presentation_delivery_report、record_presentation_issue_action、export_presentation_delivery_report。工具读取严格身份/8MiB/旧PC/clear/docchange/abort；export同PCread，仅本地VFS原子写JSON+MD（安全短路径避免超限），不切换activeArtifact、不写QA或获取复核证据授权。Controller新增readDeliveryReport/exportDeliveryReport/recordIssueAction方法，状态report只绑定当前选择request，switch/clear/docchange清理；卡片显示分组、页/主张上下文和处置表单（不默认替用户提交解释），限制一次显示问题数并明确未展示数量；旧PC保持可用。限定提交。
Root：runtime注册/路由/生命周期接线、真实plan→附件证据→复核/复算→问题处置→重启→导出跨层测试，补数值失配/恶意公式与报告读取不改变宿主/QA，文档与整体评估。独立全diff审查与最多两轮修复复审；全仓test/typecheck/lint/format/licenses及两端build。

## 退出与回滚

批次完成要求算术复现与处置/导出在真实Store/服务/插件工作台接通，不能只导出stub。所有未验证scope显式保留。JSON/MD无公式执行和自动URL请求。旧PC拒绝新操作时提示升级；新reproduction字段旧parser会拒绝，需配套升级；未使用新字段的旧plan不迁移。回退不删除已存处置记录。整体44%基线，按实际模块退出证据复评，不预设上涨。实机0/20暂缓，不merge/push/deploy。
