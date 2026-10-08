# 原方案阶段6：五类行业制作工作流

基线410e6698，接续原生图表品牌批次。原方案明确路演、汇报、培训、研究报告、销售方案行业Skills；现实现五类只有章节/问题，缺可执行制作指引及交付可见性。

采用现有read_presentation_domain_skill和原规划/生产/QA/交付工具。新增presentationDomainWorkflow(domain)仅支持五类，返回独立克隆的version1、domain、sections（id/title/instruction）、reviewSteps（id/title/tools/instruction）、manualChecks、disclosure；科研/法律/金融professionalWorkflow契约保持。无需新PC动作、依赖或数据存储。

1. Engine：presentation-plan.ts定义五类实际领域叙事与审核指引；presentation-delivery-report.ts新建报告携带domainWorkflow，严格验证出现字段的精确指导内容/领域，允许旧报告缺字段。JSON/Markdown保留全文，原输入/问题摘要、检查和专业契约保持。新tests真实生成/解析/篡改/旧报告/独立克隆，先RED。
2. Office：presentation-planning.ts返回对应domainWorkflow并指导读取、计划、生产/检查步骤；交付卡片显示报告持有的行业章节、执行步骤/中文工具名和待人工核验问题，不从现场新配置替换。更新旧返回合同测试与新React测试，先RED。
3. Root：实际PC计划保存→生产→报告读取/重启/导出链覆盖五类领域，不把章节存在当作内容真实性或行业效果通过。不同实现者交叉审查，最终完整八工作区、三端类型/静态及Office构建，记录原64%口径与待验退出。

安全与兼容：指引不能认证事实、品牌或专业语义，不产生自动处置/QA通过，不替代用户证据，不新增宿主写入。旧报告缺字段仍可读；已有professionalWorkflow无删改。回退代码提交即可，无持久迁移、外部写入或部署。工业效果退出仍须真实任务人工修改对照与实际视觉评估。
