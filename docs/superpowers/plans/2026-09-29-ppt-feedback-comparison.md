# 原方案阶段0/P2：冻结任务的人工修正反馈对照

基线7917ba3b。原方案阶段6要求行业技能显著降低人工修改，17.2定义人工修正页数。本批复用已持久生产和user_reported反馈，建立可追溯、只读的两次制作对照；不把计划domain声明、单对样本或合成反馈认证为行业效果/验收成功。

## 设计与约束

同一doc/project两次真实冻结production：baselineRequestId为domain缺省的通用计划，requestId为pitch/report/training/research/sales之一的行业计划。两次都全compiled，真实身份/planDigest/inputDigest/planRevision及反馈revision/recordedAt保留。以冻结计划的brief、sources、claims、research、style、brandKit和有效parallelism逐项比较；title/slides/domain是制作输出，不作为输入条件相等要求。计划条件相等不证明来源实际字节/模型/环境或技能实际使用一致。

报告保存两份完整冻结计划和所选评价的逐页status（不复制全部反馈历史/说明），严格浏览器安全解析并重算条件、完整评估计数和差值。服务项目锁内读取两任务和最新反馈，生成报告后下载即固定观察版本；后续评价不改已下载报告。缺反馈/未评估继续未知；只有合法通用/行业角色、计划条件相同且两边全部评价时才给描述性count/rate差值。不同页数允许展示实际分母；不输出显著性/成功/认证结论。专业science/law/finance域不冒充原五行业域。

1. A：packages/pptx-engine/src/presentation-feedback-comparison.ts、package.json导出及测试，严格纯builder/parser/types；apps/shell/src/main/presentation-service.ts注册production_feedback_compare，严格fields/requestID/doc/sameproject，existinglock内load生产与反馈，新实际PC tests。无新存储、依赖、Agent工具/能力。
2. B：Office controller和ProjectCard显式选择同项目已编译基线任务、点击读取比较、展示冻结版本/评价覆盖/条件差异及未知、完整JSON本地导出；epoch/doc/task/baseline守卫和忙状态保护，旧PC不支持明确不可用。新controller/实际DOM tests。
3. Root：真实PC保存通用/五类行业计划、编译、不同feedback、条件不一致/未评价、冻结版本变化和PC重启，加实际OfficeController到PC联调；Relay团队通道禁止该私有op。完整交叉审查最多两轮修复，fresh类型/静态、八工作区回归、sourcecommit后Office最终构建，阶段报告和严格进度。

回退代码即可，无历史文件迁移/外部部署；JSON报告source:user_reported、effect:not_verified恒定。需要真实相同任务对照、专业/宿主退出证据才可能达到原退出档位。保留64%及0/20专业现状；测试只证明对照算法和工程链，不伪造实测效果。
