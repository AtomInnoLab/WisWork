# PPT Agent：按页内容与证据预检

日期：2026-09-23。基线 `d69dd671`，分支 `codex/ppt-agent-implementation`。

## 当前进度

**整体进度约 42%（工程成熟度估算），真实专业任务验收 0/20。** 本轮首次建立统一口径；九模块等权评分及依据见[进度台账](./wiswork-ppt-agent-overall-progress-2026-09-23.md)。O4 仍为 50% 档，本轮预检未补齐来源真实性、语义和实机验收，不能按新增工具数自动抬高整体比例。

## 对应方案与完成内容

依据原方案 §11.7、O4、§21，新增内容/证据 Quality Pipeline 的确定性预检环节。原方案不变。

- 共享 `checkPresentationPageContent` 检查指定页冻结计划与 SlideIR；复用既有 schema 和计划匹配约束。
- 可见文字、表格单元格及按编译器配置显示的图表标签逐项检查字面覆盖；不把备注或跨元素拼接计为覆盖。缺少字面匹配提示人工核对，允许正确改写，不能据此断言语义遗漏。
- 列出来源摘录缺失、定位缺失、引文不在提供摘录、计算待独立复现等问题；报告包含具体页、主张、来源 ID。
- PC 新增只读 `production_content_check`，严格绑定 document/project/request/page，从指定任务冻结的 plan/deck 读取；返回计划版本及输入/计划摘要。后续修改当前计划不改变旧任务的预检依据。
- 编译前后均可预检；不启动编译、不写持久状态、不访问来源 URI、不执行公式。
- 插件新增 `check_presentation_page_content`，复用请求、取消、文档切换与生命周期保护。严格校验报告字段、身份、摘要格式、固定未核验状态和响应大小；返回固定修复建议。
- 插件预检不更换活动成果、不写 VFS/回执/QA，也不更改最近项目。旧 PC 返回明确升级提示。

报告始终保留 content=needs_review、sources/calculations/timeliness=not_verified、host=not_checked；没有问题也不自动变为通过。原宿主 QA 流程与记录格式不变，未引入全局写锁。

## 验证证据

- 全仓 `npm test` 退出码 0：6214 项 Vitest、2 项 Node 脚本检查及 53 项 Rust 测试通过，另有 6 项 Vitest 跳过。
- 独立审查发现并修复图表隐藏名称被计为可见文字的问题；修复后重跑 PPTX engine 完整测试，586 项通过（新增一项回归，包含在该数中），限定复审无剩余阻塞。
- 全仓类型检查、Lint（0 错误、9 条既有警告）、格式/相对基线格式、依赖许可与 diff 检查通过。
- Office Addin 与 Shell 构建通过；Addin 仍有大于 500 kB 的既有体积告警。
- 日志：`/tmp/wiswork-content-full-test.log`、`/tmp/wiswork-content-engine-final.log`、`/tmp/wiswork-content-typecheck.log`、`/tmp/wiswork-content-lint.log`、`/tmp/wiswork-content-build-*.log`。

定向验证包含冻结计划更新后仍读取原版本、重启、四种编译状态、跨文档/任务/页面拒绝、排队取消、磁盘文件字节不变、派生仅改变目标页预检，以及真实 runtime→PC 只读链路。checker/parser、service 和插件均有 TDD 红绿证据。没有进行真实 Office 或外部来源核验。

## 边界、部署与下一步

本次是完整内容/证据 QA 的第一层。不会获取或验证原始来源，不验证资料时效、管辖口径、数字运算、图表语义、必需内容/验收标准或共享样式依赖，也未读取用户当前宿主页文字。标题 metadata 和备注不是可见正文；图片中的文字不在字面覆盖检查范围。

报告仅针对明确选择的冻结生产任务，即使该任务在宿主上已被取代，也可作为历史产物预检；不得用它替代当前宿主验收。旧宿主保护仍由已有 QA/映射链路承担。

仅新增只读操作，无存储迁移和新增依赖；可回退插件而不丢记录。本批不合并、推送或部署。

下一步按原方案接入证据材料与主张的可追溯核验，补语义/计算/时效检查；真实 PowerPoint 环境就绪后执行保存重开、截图、可编辑性及并发验收。P0 和 O4 仍为进行中。
