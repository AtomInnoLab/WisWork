# Office 原型闭环修复计划

基线：`471b079`；隔离 worktree `/tmp/wiswork-design-draft`，分支 `codex/fix-office-prototype-completion`。用户批准范围见对应设计文档。

目标：修正既有能力的遗漏和实现错误，让模型能获取截图、登记真实验收并继续编辑含图片页面。非目标：升版、部署、PR 和 DESIGN.md 新功能。共享约束：不放宽安全/验收门槛，不记录敏感内容，不改 Relay v2 外层协议。

## 1. 截图与验收工具交付

- 文件：`packages/codex-bridge/src/tool-router.ts`（恢复清单），`apps/shell/src/main/office-codex-proxy.ts`（受限解包），`apps/office-addin/src/agent/use-office-agent.ts`（远程截图封装），按需增加纯浏览器兼容的封装/编码模块及各层测试。
- 输入：真实 `createPowerPointSkill` 工具定义与 `ToolExecution.modelContent`；输出：既有字符串结果帧、PC 恢复的模型图片块。
- 验收：Mac/非 Mac 工具清单不漏 review；无图片/无效/过大结果失败；合法截图进入 MCP 图片块；不改变已有验收拒绝条件或消息/权限大小限制。
- 顺序：缺失清单 RED 已复现，补截图链 RED，再最小修复，跑单元与链路测试。形成独立可审查改动；未授权推送。

## 2. 图文混排快照与安全错误诊断

- 文件：`apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts`、`powerpoint-skill.ts`、`apps/office-addin/src/agent/use-office-agent.ts`、现有 diagnostics 模块及对应测试。
- 验收：图片不能使普通编辑前置快照报错；文字、样式、图片几何及删除仍改变指纹；真实读取失败保留拒绝。诊断只提取安全字段，raw invalid 可区分，不泄露参数。
- 顺序：生产适配器图片复现 RED → 安全文本映射 → 诊断 RED/GREEN → 周边测试。与第 1 项共享 use-office-agent.ts 的改动由根任务协调。

## 3. 下载连接回退

- 文件：`apps/shell/src/main/office-retrieval-proxy.ts`、`apps/shell/tests/office-retrieval-proxy.test.ts` 或独立生产下载路径测试。
- 验收：预验证 IPv6/IPv4 全部交给 Node 原生回退；任一 DNS 地址不安全仍在连接前拒绝；旧 scalar lookup 行为保留；不增加 HTTP 重试或放宽预算。
- 顺序：地址丢失 RED → 固定地址列表与 autoSelectFamily → 私网/取消/重定向周边测试 → 独立审查。

## 4. 集成验证与交付

- 每个实现单元独立审查，再总览跨层图片边界与既有授权/回滚路径。修复 Critical/Important 后再验证。
- 运行 `npm test`、`npm run typecheck`、`npm run lint`、`npm run format:check`、`git diff --check`、PC Shell 和 Office Add-in 生产构建。声明真实 Mac Office 尚未执行的验证。
- 交付源码修复和证据；后续经用户明确授权，PC/Taskpane 各升一级、提交 PR 并更新本机 Taskpane。回滚仅恢复对应代码，无数据库或文稿变更。

## 实施记录（2026-09-10）

- 1–3 已实现。截图源图上限 4 MiB、模型预览上限 128 KiB，字符串再次 JSON 转义后上限 240 KiB。浏览器复用既有图像验证并生成预览，PC 先检查尺寸，再复用已接入的原生图片解码；仅原始字节和 MIME 均通过时才向模型声明可见。UI 保留原 PNG，Relay v2 无新增字段。
- 截图准备或封装失败、取消、期间文稿修改或契约替换，均不得清除较新的待检查状态。远程处理器在发布活动前统一处理封装失败，UI 和本地诊断不再虚报成功。
- 补回验收工具清单；混排快照保留全部形状几何及可用文字/样式指纹；下载保留全部已验证公网地址并使用原生双栈回退。
- Office 错误链只提取有界标识，另排除 URL、路径和 IP 形式。raw 输入补充现有精确字段说明，失败仍在宿主操作前拒绝；不改变解析器或权限。
- 所有可复现路径均补充 RED/GREEN 测试。独立审查发现的无效图片提升、过期截图清除新修改状态、封装失败 UI 误报和诊断字段边界已修复并复审；最终独立审查 387 项通过，无剩余发现。下载和快照另有独立审查通过。

### 验证证据

- 真实插件处理器 → Relay v2 输出帧 → PC 代理 → MCP HTTP 图片块链路测试通过（宿主截图为固定测试图，不是真实 Office）。
- 实际 Codex 0.147.0 执行器配本地假服务：5 项工具循环集成测试通过，9 项未选中；不是付费模型验收。
- 本地 Chromium 原生 `createImageBitmap`/canvas 验证：1920×1080、2,744,596 字节及 1440×810、1,550,297 字节 PNG 分别变为 1280×720 的 57,618 / 76,322 字节 JPEG；原图不变，超限源图拒绝。
- PC Shell 和 Office Add-in 生产构建通过；仅既有 bundle 体积等构建警告。未打包或发布 PC 安装包。
- 最终 `npm test` 退出 0：7,498 项 Vitest 测试通过，21 项既有条件测试跳过；仓库要求的运行时策略、执行器发布与 macOS 打包脚本检查均通过。Office Add-in 最终 894 项通过。
- 最终 `npm run typecheck`、`npm run lint`、`npm run format:check`、`git diff --check` 全部退出 0；lint 保留既有 13 条警告、0 错误。
- 原始全仓测试曾读到并行补充的过期截图 RED 用例而失败；修复、独立复审后已完整重跑，不以先前部分通过替代最终结果。

### 交付边界

现场诊断已脱敏，无法证明每个 `image_fetch_unavailable` 都由 IPv6 连接失败引起，或每个 raw invalid 的实际参数是什么。不能把剩余外部图片源故障归为已解决；本次修复确定的代码缺陷并补足诊断。尚未在用户的 Mac PowerPoint 上实机验收，也未编辑、删除或覆盖其半成品文稿。

按后续授权从隔离分支提交并提 PR，更新本机 Taskpane 部署；不合并 PR、不发布 PC 安装包、不部署 Relay。未纳入原有的 DESIGN.md 同步设计文档或其他无关改动。
