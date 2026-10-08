# 独立研究账本的持久语义时间线

基线33957772。按原方案§7.1/7.2，使用真实既有ResearchStore归档投影，避免摘录核对冒充独立研究，不新增事件存储或自动研究操作。

## 单元与接口

1. A：PC presentation-service status在既有项目锁内读取researchStore.summary，增加可选researchSummary；异常researchHistoryUnavailable:true，保留项目。使用原V1/V2最近32摘要，不含草稿。实际begin/finish/结束/删除/重启/损坏测试。
2. B：Office presentation-project严格共享解析可选摘要，校验project及当前document，unavailable优先，不保留未知字段；旧status兼容。独立控制器测试。
3. C：workflow添加scope research_ledger稳定记录身份事件与真实开始/终态子记录；显示来源/主张/冲突计数、精确当前绑定、窗口/失败/未决提示。记录不证明事实认证或运行中，不改变QA完成与生产动作。摘录核对改source_audit.*，保留原scope。
4. Root：更新旧摘录事件断言；实际PC/pairedRuntime/Office控制器跨层验证归档、终态、重启、删除gap及冻结原绑定不被后来记录替代。

各单元先RED后GREEN，互相独立审查，root统一相关回归、四端类型、静态、源码提交后Office构建及阶段文档提交。保持64%成熟度口径，真实专业任务0/20，不部署。

## 执行结果

A/B/C/Root已完成及交叉独立审查通过。最终完整相关2949/2949、定向91/91、展开2/2、四端类型/静态通过；源码56e5495a，Office构建56e5495a7ece。首轮工作台测试失败及idle前置条件修正见本批报告。总体64%，真实任务0/20，未部署。
