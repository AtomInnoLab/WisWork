# 几何保存点检查与确认恢复

基线48d6d76c。依据原方案§6.4、§8.7、§11.8、§14、§20/21；高保证/TDD/独立审查，复用已有状态机、几何adapter与proposal。

## 范围/算法

新增inspect_presentation_geometry_change与resume_presentation_geometry_change，inputs project_id?、page_id必需（resume可explanation）。只针对pending/undo_pending；读取当前宿主geometry并严格校验身份、artifact/receipt/日志一致。

pending：origin=before,target=after；undo_pending：origin=after,target=before。容差0.01，只有匹配target→already_applied，只有匹配origin→ready_to_apply，双方匹配或均不匹配→manual_review。inspect只回读，不写日志/宿主，不解释历史因果。终态返回not_pending而不是宣称当前正确。

resume需新确认；manual_review不能propose。validate/execute重新回读并核对相同classification与精确观测值，不允许确认后状态漂移。ready_to_apply向既有adapter传target和刚回读actual，已有pending就是预写记录，无新状态。already_applied不重写对象。verify回读target后CAS保存applied或undone，失败保留pending。每await后check身份/epoch/日志/取消；写后verify不被Stop跳过。无新record/schema/设置项。

frontend负责page-editing.ts+tests，独立TDD与scopedcommit，覆盖两类pending的两条恢复分支、both/neither拒绝、宿主改变/取消/文档切换/保存失败/重复恢复与0.01量化。旧undo/read行为不变。
root负责runtime路由、resume可信局部QA白名单、跨runtime+实际DocumentBinding恢复集成测试、阶段报告与验收清单。reviewer独立检查。全仓npm test/typecheck、变更Lint/diffcheck；测试结束后Addin/Shell构建。不实现文本/图片/整页撤销或单页重做；实机Office/20任务仍待执行。无数据迁移、PC/Relay/manifest变化，回滚保留不确定日志。
