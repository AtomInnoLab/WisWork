# 按页多来源复核历史汇总

基线0f29babf，按原方案§4.3/O4；用户要求继续但暂不实机验收。整体成熟度42%、实机0/20暂缓。高保证/TDD/独立review/全仓验证，原方案不改。

## 设计与接口

新增只读production_page_reviews(operation/documentId/projectId/requestId/pageId全必需)，按指定冻结production的page.claimIds→claim.sourceIds顺序汇总已有不可变review。新增store.listClaimReviews(projectId,documentId,requestId)复用private claimReviews全验证，缺任务返回[]（service先验证production存在），返回clone避免调用者改内部状态，无新存储/迁移。

共享module presentation-page-reviews.ts export ./presentation-page-reviews。PresentationPageReviews字段：{version:1,projectId,requestId,pageId,planRevision,inputDigest,planDigest,claims:[{claimId,status,sources:[{sourceId,status,reviews:[{reviewId,outcome,evidenceDigest,createdAt,offset,maxChars}]}]}],checks:{support:'historical_agent_reviews',sourceAuthority:'not_verified',timeliness:'not_verified',host:'not_checked'}}。review outcome固定supported/contradicted/insufficient_evidence。

source status：无review unreviewed，所有outcome同一个用该值，否则mixed。claim status：无source no_sources；所有source unreviewed则unreviewed；任一mixed或已评source的不同outcome>1则mixed；其余有unreviewed则partial；其余用共同outcome。statuses仅描述历史Agent判断分布，不是真实事实通过。不采用last-write-wins，也不把不同窗口/不同时刻的判断合成客观矛盾。历史仍全部保留，mixed需要人工/Agent回看原记录。

最多32claims，每claim3sources，总review<=32且reviewId全局唯一。ref不含notes/原文，报告<=64KiB。严格parser验证字段/枚举/ID/摘要/时间/范围/唯一性并重新计算source/claim status拒绝伪造。输出有完整源列表，未评来源不能被省略。实现helper summarizePresentationPageReviews(plan,deck,metadata,reviews)接口可agent自定，parser名字parsePresentationPageReviews固定。服务必须验证选中page的store review与冻结claim/source归属，不能静默忽略不合法引用；其它page正常不纳入。

插件新增read_presentation_page_reviews(project_id/request_id/page_id)。复用production skill只读request/current/cancel，严格parser+身份+64KiB，结果不切换artifact、不rememberProject、不写VFS/QA、不创建写复核授权；提示可用read_presentation_claim_review按reviewId看理由，read_presentation_claim_evidence重新读取后判断。旧PC升级提示。

## 分工验证

backend agent owns sharedmodule/tests/export、store列表方法/test、PC service/production/newtests。root owns插件/runtime/单元及真实integration扩展、阶段报告。独立review后全仓test/typecheck/lint/format/licenses，再buildAddin/Shell。TDD覆盖缺失来源、mixed/partial、无来源、旧review不覆盖、不同页隔离、冻结plan更新、重启/坏记录/跨doc/取消及只读。不做实机验收。

## 边界与回滚

只读新增接口，无磁盘迁移、不自动重新抓取附件、不自动判定过期、不修改/合并旧review或将其变成当前宿主验收。不新增全局门禁，下一步补复核变更处理或计算/时效可追溯检查。不合并/推送/部署。

## 完成记录

共享契约、存储查询、PC 服务、插件路由及跨层测试均已实现。独立审查与全仓测试、类型检查、lint、格式、许可证、两端构建通过；详见阶段报告 `docs/product/wiswork-ppt-agent-page-reviews-progress-2026-09-24.md`。整体成熟度仍为 42%，实机验收 0/20，按用户要求暂缓。
