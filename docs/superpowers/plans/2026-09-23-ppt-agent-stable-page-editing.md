# PPT Agent 按业务页定位与文本原位修改

基线1de16e41，继续已批准方案的Office.js原位修改。沿用隔离分支、TDD、独立审查与完整验证，不合并/推送/部署。

## 目标与边界

业务页ID通过已确认导入检查点映射为宿主页ID，读取原生对象并提出单对象文本修改。所有新路径直接通过宿主页ID访问Office对象，页面重排不改变目标；缺失/被替换页不猜测重绑定。不新增图表/布局/图片修改、单页重生成或撤销，不声称实机验收。

## 单元与接口

1. browser-powerpoint-adapter.ts与新测试：新增listPresentationPageShapes(slideId,signal) => {slideId,shapes,shapesTruncated}（最多100，读取101识别截断）；readPresentationPageText(slideId,shapeId,signal)=>SlideTextResult；editPresentationPageText(slideId,shapeId,text,expectedText,signal)=>void。复用现有文本写入/回读/取消归因代码，以getItem(hostId)直接访问，写入前比较expectedText，缺失或不一致拒绝。新方法Office1.10能力门控；旧index方法行为不变。合法ID<=256、文本<=12000，不混入业务ID。RED/GREEN后scopedcommit。
2. 新presentation-page-editing.ts及测试：两个工具read_presentation_page(project_id?,page_id,shape_id?)和edit_presentation_page_text(project_id?,page_id,shape_id,text,explanation?)；前者无shape读对象列表，有shape读文本，后者走现有StructuredProposalController。配置available/artifact/documentId/readReceipt/adapter/proposals。校验文档、artifact身份/源包SHA256、源页顺序、已完成页映射；每次异步后复核上下文；validate与execute再次检查映射/原文本，不以index定位。clear使旧提案失效；输入严格有界、未知字段拒绝，输出最多64KiB。旧legacy无检查点不猜测绑定。RED/GREEN后scopedcommit。
3. root runtime接工具、清理生命周期、提示默认使用业务页工具；跨层测试验证重排只修改同一host页、删除目标不写、修改后QA标记及新截图复核、恢复/取消/陈旧提案。阶段报告与提交。

## 安全与回滚

保留现有用户确认、修改前QA失效、写入前二次校验与回读。Office没有原子CAS，记录预读到sync期间仍有宿主竞态，不自动回滚覆盖用户新修改。工具/接口只做增量，无持久化格式变化；回退新工具即可，已修改文稿不自动逆转。

## 验证

独立审查全diff，目标回归、全仓test/typecheck、变更lint、Office插件与Shell生产构建。Office依赖以mock测试，真实PowerPoint重排/删除/保存重开留作实机验收。
