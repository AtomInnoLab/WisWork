# PPT Agent：共享样式依赖与局部重审

基线08893628，既有隔离实现分支。原方案§5阶段F、§6修改流程、O4要求共享样式影响页重审且不误伤无关页。沿用用户已批准方案，实机验收暂缓。高保证：TDD、子代理实施、独立审查、全仓验证；不改原方案。

## 架构与边界

直接读取Office页面的slideMaster.id/layout.id建立瞬时依赖快照。原生母版背景/主题色覆盖引用该母版的页面，版式继承只覆盖该母版下的目标版式；不推断颜色相等即依赖，不尝试判定对象级主题覆盖。提案保存完整快照并在确认校验（含持久失效后的二次校验）重读；变化则阻止写入。依赖不可读/不完整/超限时明确退回文档级QA失效，XML和未知脚本仍全量失效。局部对象编辑仍只影响本页。

复用现有recheckRequired和串行Office settings保存，不另存依赖图、不迁移旧记录。空的已知依赖范围表示没有页面受影响，undefined表示未知范围且全量失效；两者严格区分。继续在写入前持久化失效，失败禁止宿主写入，回滚后保留失效保守标记。不会自动将重采集判为视觉通过。

## 契约

新增presentation-style-dependencies.ts（插件browser-safe）：

- PowerPointStyleDependencies={slides:Array<{slideId:string,masterId:string,layoutId:string}>}，最多100项，三个ID非空<=256且无控制字符，页面ID唯一，严格字段；完整性由原生读取全部slides保证，超限抛错。
- parsePowerPointStyleDependencies(unknown)返回校验副本，排序按slideId；affectedStyleSlideIds(snapshot,operations)返回去重排序宿主页ID。
- PowerPointAdapter.inspectStyleDependencies?(signal):Promise<PowerPointStyleDependencies>，Browser实现可用；旧adapter不实现时降级。
- 原生edit_slide_master提案preview.qaScope：已知为{basis:'native_master_layout',hostSlideIds:string[]}；未知为{basis:'document'}。外部工具输入不允许直接提供qaScope。范围完全来自已验证宿主快照与内部解析的operations。
- 已知依赖提案validate必须重读并严格比较快照，读失败返回false或抛错而不降级继续；初始读取未知则维持全量范围。signal取消不能吞掉后继续提案。
- Root beforeWrite只信任operation/toolName都为edit_slide_master且hostpowerpoint的严格qaScope；其余既有单页白名单不变。presentationQaMutationScope支持合法空数组，保留对非法/超限/重复/控制字符的拒绝。beginMutation([])锁定操作但不删除无关live截图；invalidateQa([])无写。

## 交付单元

A native依赖（agent）：新dependency模块/tests，browser-powerpoint-adapter.ts及适配器tests，powerpoint-skill.ts及原生mastertests。先RED：不同母版/同母版不同版式、主题/母版/版式作用域、空页、未知/重复/超限/取消、两次validate快照变化阻止写。实现完整快照与提案范围、保留现有回滚。提交限定文件。

B工作台重审（agent）：presentation-qa-card.tsx、App.tsx与卡片/App测试；保留controller接口，通过可选onRecheck(record,pageIds)回调把明确project/request/pageIds重审指令填入composer，不直接执行、不自动pass、不绕过proposal。新增受影响页数量/列表、单页及受影响页批量准备重审按钮，busy/applying/proposal/upload时禁用。不把任意标题/notes作为指令注入，只引用校验ID，指令明确先确认目标冻结任务与宿主映射、逐页截图并观察后复核。历史结果仍标注限制。独立测试点击范围/禁用/未配置回调/无法读取/历史pass优先失效。提交限定文件。

Root：host-runtime beforeWrite范围接线；mutation scope空集合语义；document持久失效与现场截图授权范围跨层测试；方案/阶段报告与进度台账。用真实binding、真实proposal生命周期、模拟Office截图和依赖，证明母版/版式受影响页失效、无关页记录/live授权保留、重启后状态、持久失败和二次validate阻止写、未知范围全量及空范围无写。先RED后实现。

独立review：全diff一次，最多两轮修复复审，重点完整性、TOCTOU、fallback与空范围、Agent截图授权、UI提案/文档身份和指令不自动放行。完成后全仓test/typecheck/lint/licenses/format/diff，测试结束再两端build。

## 回滚、兼容与进度

无外部依赖、无持久格式变化；旧adapter无方法自动全量失效；旧数据可直接读取。恢复旧版本只退回保守全量失效，不删除任何历史记录。Office能力真实Win/Mac/Web验收仍未做；本批只验证工程行为，真实专业任务0/20。整体44%基线，按既定九模块口径评估，不按提交数量上涨。仅实现分支提交，不merge/push/deploy。
