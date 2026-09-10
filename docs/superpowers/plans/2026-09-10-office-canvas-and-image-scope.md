# Office 实际画布与插图页面归属修复

基线：`252c13c`，分支 `codex/fix-office-canvas-and-image-review`。
用户授权：定位半成品错误、修复、PC 升版并部署 Taskpane。保留提案确认、原型、截图和几何验收；不修改用户文稿，不改变 Relay 协议，不纳入 DESIGN.md 同步草稿。

## 证据

- 最新 trace `6608de63-4946-4e4a-b8ad-4694defba226` 使用已部署 build `826686604fd8`。图片和文字整体只占左上约 75% 区域，与使用错误的画布尺寸一致；无 PPTX，尚不能断言原始文件具体尺寸。
- `get_presentation_state` 没有返回实际页面尺寸，只有末尾 `verify_slides` 才读取 `pageSetup`。Office 的 slideWidth/slideHeight 单位为点，不是截图像素。参见 [Microsoft PageSetup 文档](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.interfaces.pagesetupdata?view=powerpoint-js-preview)。
- PowerPoint skill 用共享 `proposingDesignSlides` 临时变量给全局提案审计归属页面；插图在另一 skill 提案，误继承上一次文字修改的页码。真正插图页仍被当作新生产页，背景修复遭待审门槛拦截。
- 图片写入未接入同一生产前置检查。应同时修复归属和门槛，而不是允许无限扩页。
- `design_contract_visual_review_failed` 留在本地诊断，但漏了旧 Relay 错误码映射；服务端白名单拒收，产生次生上传失败。
- 字体回读失败定位到了 fontFamily，但公开读文本工具没有提供现有字体样式；应提供可读取的真实样式帮助局部修复，不接受字体未生效为成功。
- 图片 15 秒失败仍缺少 URL 和网络阶段证据，不宣称任意外部图片源已恢复。

## 实施与验证顺序

1. 写失败回归：真实尺寸在制作前可用；插图不能污染上一次文字页的验收；视觉错误兼容旧 Relay；字体样式可读。
2. get_state 返回可测量的点尺寸并保留上下文，无 API 时不伪造默认画布。明确整页边界、图文对比和图片比例要求。
3. 每个提案携带内部、限长、复制后的页面归属；独立图片模块使用自己的页码并走生产前置检查。用户输入不能注入该内部字段，Relay 格式不变。
4. 兼容映射诊断码；保留本地精确失败。文本读取附带可用的真实样式；不支持时仍保留文本结果。
5. 定向回归、独立审查、全仓测试、类型、lint、格式和生产构建。升 PC 0.6.66 / Taskpane 0.3.36.0，提交 PR；备份本机静态站点，先资源后入口替换，并从公开 HTTPS 比对构建产物。

## 发布边界

无需 Relay 服务升级。无数据库迁移；保留旧哈希资源及部署前备份供回滚。源码版本及 PR 不代表用户已安装 PC 新版。自动化测试不代替真实 Mac PowerPoint 的全流程视觉验收；新修复不会自动重做现有半成品。

## 验证记录

- RED 已复现：缺失实际点尺寸、跨模块插图误记到上一文字页、尺寸未保留上下文、视觉错误码未映射、读文本不返回实际字体样式。修复后回归转绿。
- `npm test` 退出 0：Vitest 7,518 项通过，21 项条件测试跳过；Office 插件 908 项通过。
- 全仓 `npm run typecheck`、`npm run lint`、`npm run format:check`、`git diff --check` 均通过。lint 为既有 13 条警告、0 错误。
- PC Shell 生产构建通过；独立代码审查及字体读取/升版增量复审均无阻塞项。
- Taskpane 将在提交后使用提交哈希构建生产资源；部署备份和公开 HTTPS 比对结果记录于 PR。
