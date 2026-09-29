# 原生母版持久恢复实施计划

原方案§6.4/§14.2；已核对当前edit_slide_master的实际能力。此次XML前像守卫不代替本计划。

## 保留合同

version2程序1–32个不重复目标key；多master/layout混用；solid/gradient/pattern/picture_or_texture背景、12个theme颜色slot、layout继承与母版图形开关。读取32个master、每master128个layout。现有不可逆picture原背景拒绝条件保留，不额外缩小已接受操作。

## 专有账本与执行

账本独立于原页替换的v1–v4 union，因为恢复页包不能撤销原live master/theme/layout。持久原目标值、期望值、完整依赖、原页序及原始包保存点，operation及inverseOperation、cursor/pendingIndex、actual receipts/state。复用当前parseMasterProgram/projectedMasterState/masterOperationValue/inverseMasterOperation/affectedMasterFingerprint。图片改用PC内容摘要引用，避免原base64撑爆settings。新记录browser安全严格验证和CAS转移，预算必须容纳32操作及完整受影响依赖，不套8页批次限制。

1. 捕获实际ID及完整依赖。当前inspectStyleDependencies超过100页明确失败；不得将失败等同零页，也不得取前20视觉页做备份证明。完整依赖能力及预算先落实。
2. 保存并读验全部必要原包保存点，重新核文档/目标值/依赖/原页序，保存账本。
3. 每项写前持久pendingIndex；只调用executeMasterOperations([op])；readback确认后保存receipt和cursor。异常/取消保留未决，不执行旧catch自动逆写。
4. 只读inspect分类目标before/after/unknown，unknown禁止继续；明确确认恢复丢失回执才推进，不盲目重放。
5. 撤销提案核所有已写目标及依赖，按逆序执行inverse，每项同样先intent再readback回执；外部修改不覆盖。
6. Runtime binding/lifecycle/history/workbench专有kind及inspect/reconcile/undo路径；不派发到恢复原页替换。QA范围依赖实际master/layout影响集，历史截图不认证完成。

## 验证与后续

合成实际PC备份+document settings+native SDK，覆盖多母版布局、32操作、完整依赖超100页、不确定ACK、重开、逆向撤销中断、外部漂移、取消/断连/ID切换、图片预算及CAS。独立审查、八工作区回归及生产构建；真实宿主另按原退出条件验收。XML母版导入创建新master且对原master关联页面applyLayout，需要另行持久原layout/newlayout映射和逐页回执，不能只恢复第一页。
