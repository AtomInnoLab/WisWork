# 持久宿主页面导入语义（2026-09-29）

基线3c273938，原方案阶段1/阶段5的checkpoint、语义进度与重开恢复。现有inFlight.startedAt成功时丢失；用可选startedAt延伸既有completed页记录，真实开始时间与完成记录时间分别保留。旧记录不补猜；重核认领保留原开始，认领时间不是原写入完成时间。回执身份、确认、插入、QA和恢复策略沿用既有实现。

## 单元

1. A：apps/office-addin/src/skills/powerpoint/presentation-delivery.ts、presentation-page-delivery.ts及page-delivery测试。成功/只读重核后保留原inFlight开始；严格时间戳、先后顺序和字段校验，旧无时间形状兼容，状态投影保留双时间。先失败回归涵盖实际导入、回执丢失认领、旧记录和非法时间；不改变写入/撤销权限或实际宿主行为。
2. B：apps/office-addin/src/agent/presentation-workflow.ts及workflow/card测试。从当前精确project/request匹配的真实progress派生host.import.started/recorded/uncertain，scope=host_page_import；默认折叠双时间及已记录/待核查说明，身份绑定任务/页/宿主slide。旧缺开始不造事件，完成仍非QA完成，任务/页面状态与nextTool保留；不混入别任务或当前计划。
3. Root：跨真实skill/receipt/workflow重新读取的集成回归，基于真实PptxGenJS编译与可控宿主（非实机）证明正常写入及回执丢失核对后的时序/身份/旧记录；独立C只读审A/B完整diff，修复后复审。

## 验证及交付

各单元按TDD先RED后GREEN，定向周边回归/类型/静态；Root统一完整Office及相关Shell/engine/store/harness回归和生产构建，源码及阶段记录本地提交，未部署或上传。共享工作树隔离已有，所有代码操作仅实施分支，agent不commit/全量/build或改别人文件。可选字段不更改既有schema版本；回滚源码后新版含startedAt记录旧客户端可能拒绝，必须保留原件而不是移除时间数据，混版本待实机验证。新字段仅时间元数据、不含用户内容；不从时序认证实际PPT/QA/专业任务通过。整体64%按原成熟度档保持，真实0/20，材料17/20。

## 本批结果

A/B实际runtime契约与时间线完成，root真实编译+持久binding跨层59/59，独立审查同59/59；minor界面点击5/5通过。完整相关3549/3549、Office/PC类型/8文件静态通过，源码f6402ff7，Office构建10.09s通过，未部署。旧fixture真实无时间形状更新；实际宿主ID错误在先失败联调修复。真实PowerPoint/混合版本仍未验收，总体64%、材料17/20、真实专业0/20。
