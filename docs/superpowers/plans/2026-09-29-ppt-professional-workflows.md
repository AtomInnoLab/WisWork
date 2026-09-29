# 科研、法律、金融制作技能与工作台

依据原方案 §4.3.1、专业领域技能与研究→规划→制作→验收流程；基线072a4ad0，已有隔离实现分支。高保证、独立实现和交叉审查；不增加外部provider或真假专业认证。

## 固定接口

engine presentation-plan.ts 增加 science/law/finance 三种 domain profiles，各5必需 sections：
science: research_question, methods_and_sample, results_and_data, scope_and_limitations, research_references。
law: legal_question, applicable_materials, analysis_and_alternatives, legal_risks, legal_conclusion。
finance: financial_question, reporting_basis, financial_results, financial_risks_and_scenarios, financial_actions。
保留原research与所有旧profile字段、sections和行为，章节是组织约束而非内容真实性证明。

engine同文件导出 presentationProfessionalWorkflow(domain?:string)，未知/旧domain返回undefined；匹配三领域返回readonly结构 {domain:'science'|'law'|'finance',sourcePriority:string[],contextFields:string[],reviewSteps:{id:string,title:string,tools:string[],instruction:string}[],manualChecks:string[],disclosure:string}。内容须按原表完整优先来源/专业上下文/自动警告和语义检查范围，说明字段完整或执行流程不认证事实；拒猜时点/法域，unknown保留，来源双方/限定不为叙事裁剪。tools只引用实际存在 read_research_ledger, save_presentation_plan, read_presentation_claim_evidence, record_presentation_claim_review, check_presentation_page_content, read_presentation_delivery_report，可添加原已有tool但先核对真实runtime。

read_presentation_domain_skill专业domain返回原通用 {domain,title,sections,labels,questions} +professionalWorkflow；旧5domain outputexact unchanged。修改Agent规划prompt以科研/法律/金融任务先读对应domain技能、绑定原研究，设置domain/对应章节，未完成manualChecks不宣称验收成功。

## 单元划分

A（remote_acquisition_pc）：engine presentation-plan.ts profiles/helper及新tests、Office presentation-planning.ts工具输出/prompt及tests。坚持章schema和runtime同步，profilehelpers返回clone而不是泄露可修改全局对象。只读tool不应PC调用或写入。正式domainPlan保存/冻结须实际验证每章必需、跨domainsection拒。

B（addin_build_version）：engine delivery-report.ts专业workflow无professionalContext的claim新增 professional_context_missing（unverifiable）。每claim/code≤1，所有声明主张适用；已有专业context可含跨领域支撑事实，不误判为域不匹配。digest覆盖workflow定义+完整claim+planDigest（必须体现规则变化）；MD输出完整workflow指引及manualChecks未认证。专业workflow页cap1312=1280+32；专业context而无workflow维持1280，其他旧1056/608不变，8MiB报告不变。tests包括missing/supplied/跨领域allowed、partialcontext仍incomplete、处置stale/open和旧兼容。

C（remote_acquisition_office）：Card profesional missing原因中文，默认fold显示所选技能sourcePriority/contextFields/reviewSteps/manualChecks/disclosure，与actualsnapshot domain绑定，report更新切换domain无陈旧指引；有专业workflow却该claim无professionalContext显示缺失与下一读原文动作（不要虚假认证）；原acciones身份/cas保持。真实click和旧无workflowhide测试。

Root：实际Runtime+PC跨层验证8页保存/冻结science/law/finance各一次，先读local专业工具内容与实际tools存在，再读完整证据/报告/重开依旧domain；缺专业context的领域任务生成missingissue并持久处置，不因supported/解释关闭。改domain章节缺失 save拒不覆盖原计划；交付JSON/MD含精确workflow。更新阶段报告/台账，完整相关回归/三端type/lint+format/diff、交叉审查≤2轮、统一sourcecommit/build+buildId/docscommit，无上传部署。

专业语义（统计范围/图表原数据、限定语法效力、口径可比/预测依据）仍需实际模型和专业任务门槛；本批是实际可读取技能/编排/警告链，整体成熟度64不虚报100。

## 实施结果

A/B/C/Root工程单元完成并交叉审查；报告额外保存实际professionalWorkflow version1全文，专业必有/旧字段必须absent，冻结任务领域选择与当前规则区别已明确。full2826/2826、final23/23、三端types/static通过；补窄范围字段中文UX修正（Card6/6、Office types/static、独立复审通过），源码9eae9f80。总体64%，manual专业验证及真实宿主仍未退出。
