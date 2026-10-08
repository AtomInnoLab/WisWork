# 持久QA截图、复核与失效语义

基线20d65f11；原方案阶段4/5要求局部失效与PC/插件可恢复阶段进度。现有QA持久页保存capturedAt/reviewedAt/recheckRequired，失效发生时间缺失。新增可选invalidatedAt记录第一次真实由未失效变成失效的时间；重复失效保持原值，旧已失效无时间不补造，重新截图创建新证据并清除旧失效元数据。

1. A仅presentation-qa.ts、presentation-document.ts和qa-binding/qa测试：严格invalidatedAt必须有recheckRequired=true且规范ISO、>=capture/review；固定元数据额度延伸既有recheck字段预算，64KiB内容额度及settings总量不缩水。document.invalidateQa串行锁内首次失效保存clamp(now,capture,review)，已有stale页不重写时间，未知范围仍整稿；所有scope/保存失败/文档隔离门禁保持。测试先RED后GREEN，旧满额度与重开/重试/idempotent/recapture等必要验证。
2. B仅presentation-workflow.ts及workflow测试、新QA timeline/card测试：匹配当前生产项目/任务页的真实Qa记录，稳定identity doc/project/request/page/host/capture/digest，scope saved_page_qa，父qa.capture.recorded/qa.visual.recorded/qa.evidence.invalidated，折叠typed子事件和真实时间；原始capture、agent复核及首次失效有分别身份，fallback明确宿主外观未验。旧stale没有时间仅说明不造invalidation event，全部QA/attention/stages/nextTool现有逻辑保持。
3. Root真实QaSkill+settings重开/作用域失效+workflow跨层集成，实际编译/原生页包/受控宿主和PNG；C独立完整只读审所有单元、修复后复审。

TDD/独立审查/完整相关回归、Office/PC类型、静态和实际Office构建；root统一本地提交+阶段记录，不部署/上传。所有mutations实施worktree require_escalated；agent不commit/full/build或改别人文件。可选字段不改版本，旧端可能拒绝新metadata，保留原记录、混版本实机待验；64%/材料17/20/真实专业0/20不因tests或事件数字升档。QA记录只是历史截图与Agent意见，不认证专业事实、当前真实宿主外观或RoundTrip。回滚不能删除已有证据时间。

## 执行结果

A持久层51/51、B周边38/38、Root最终6文件85/85；独立A51/51及最终85/85无严重/重要发现。完整相关3559/3559（299文件，83.90s）、Office/PC新鲜类型、8文件静态检查通过。源码dbc0cd5a，构建dbc0cd5af228成功（9.76s），未部署。严格64%、17/20候选、0/20真实专业保持。invalidatedAt仅首次持久标记时间；QA只是每页最新记录，失败尝试和全截图追加归档不在本批。详见本轮QA阶段报告。
