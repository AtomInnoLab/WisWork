# 专业结论范围复核与交付归档事件

基线5a58315b；按原方案§4.3.1、阶段G与§7.1。沿用sourceAssessment不可变历史及原交付包回执，保持普通生成不新增确认门槛，不部署。

## 固定合同

sourceAssessment可选professional={context:PresentationProfessionalContext,checks:[{aspect,outcome,reason}]}。science/law恰好conclusion_scope与qualifications；finance恰好comparability与forecast。每项outcome为consistent/conflict/uncertain/not_applicable，最后一种仅forecast。reason1..600，不重复aspect。非uncertain需已有basis非空，全部原文basis按原真实读取UTF-16窗口逐字核对。context必须与冻结claim的完整专业上下文canonical一致，不猜域、不删限定；无真实上下文拒绝。16KiB边界及旧无专业字段契约保持。所有评估是Agent历史判断，不认证事实/权威/时效。

## 单元与ownership

1. A：共享source-assessment schema/parser/新contextassert helper及PC production_record_claim_review实际原归档/窗口写前校验，专用tests。默认无professional不新增字段，存储沿用完整immutable review。
2. B：Office production tool已有schema引用、实际seen.evidence上下文校验、提示语和精确响应nested对照，专用tests。无新tool，原read-before-write与doc/epoch保护不变。
3. C：报告按所有专业历史意见产生四aspect的conflict/uncertain/mixed问题与完整原文/上下文摘要，positive不关闭旧问题；仅新评估存在扩cap256。strictparser匹配冻结上下文，JSON/MD及默认折叠中文界面保留历史和非认证说明，专用tests。
4. Root：实际PC→Agent→报告→JSON/Markdown/当前宿主ZIP验证science/law/finance与原冻结版本、不一致专业记录拒绝、说明不关闭问题；另从原deliveryBundles回执给workflow添加delivery.bundle.started/ready事件、真实上传/完成时间和未验收说明，workflowCard接实际snapshot并专用tests。ready不生成project.completed。

交叉独立审查：A看B/root，B看C，C看A/root。各单元RED→GREEN后source freeze；root统一完整相关回归、四端types/static/diff、源码提交后Office构建、阶段文档提交。总体64%成熟度口径，候选14/20，真实任务0/20。

## 执行结果

A/B/C/Root完成并交叉独立审查通过。额外修复bundleNotice同时包含成功提示不能代表history unavailable：新snapshot flag仅实际读取失败置位，实际成功清除；反例先失败后通过。Root实际三域跨层34/34，完整相关2980/2980，交付时间线与组件5/5，四端types/static通过。源码3f7889e1，Office构建3f7889e11703；总体64%，真实任务0/20，未部署。
