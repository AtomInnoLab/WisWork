# WisWork 集成实施计划

日期：2026-08-05
设计依据：`docs/superpowers/specs/2026-08-05-wiswork-integration-design.md`
分支：`feature/wiswork-integration`
实施起点：`2234f41`

## 目标与非目标

目标是将仓库的产品身份、认证和默认模型链路完整切换到 WisWork、Logto 与 WisModel，同时保持现有本地 Office 编辑能力和 Agent Loop 行为。

本计划不实现 WisUsage 计费、Gateway 模型转发、正式生产服务 Key 分发，也不替代暂时下线的 Genspark 图片生成、媒体分析、云端幻灯片生成和 PDF 云转换。

## 架构

`@wiswork/auth` 管理 OAuth 协议、会话状态和 Electron 主进程适配；renderer 只通过 typed IPC 获取非敏感状态。`@wiswork/ai-provider` 的 `wiswork` provider 复用 OpenAI-compatible streaming/chat 实现，服务 Key 只在主进程从 `WISWORK_MODEL_API_KEY` 读取。统一 Shell 注册深链并向所有编辑器提供共享登录状态；standalone 入口使用同一认证 bootstrap。

## 全局约束

- 不在任何文件、提交、日志、快照或 renderer IPC 中保存用户提供的模型 Key、refresh fixed code、access token 或 refresh token。
- OAuth 回调必须验证 scheme、host/path、state、过期和单次消费；不得用宽松解析绕过未知后端字段。
- `safeStorage` 不可用时拒绝持久化，不降级为明文。
- 模型 base URL 与 Authorization header 由主进程决定，renderer 不可覆盖。
- 所有行为变更遵循 RED → GREEN；机械改名使用确定性扫描和 typecheck 作为验证。
- 每个任务形成一个可独立审查的提交，并由独立 reviewer 审查 Critical/Important 问题。
- 旧 userData 只复制，不删除或覆盖。
- 保留的 `GenOffice`/`Genspark` 字符串仅限迁移兼容、历史测试夹具或设计文档，并通过允许列表说明。

## 主要文件职责

- 根 `package.json`、`package-lock.json`、16 个现有 workspace `package.json`：npm scope、workspace 依赖和脚本。
- `packages/auth/**`：OAuth 配置、PKCE/state、回调、刷新、会话与 Electron 安全存储。
- `packages/ai-provider/src/{types,providers,stream,chat,index}.ts`：WisWork provider 与 OpenAI-compatible 模型调用。
- `apps/{docs,sheets,slides}/src/main/**`：共享/standalone AI IPC 与认证门禁。
- `apps/shell/src/main/index.ts`：统一进程 OAuth 协议注册、深链路由、账户 IPC 和旧数据迁移。
- `apps/{docs,sheets,slides}/src/{preload,shared,renderer}/**` 与 `apps/shell/src/{preload,shared,renderer}/**`：非敏感账户状态和 WisWork 登录 UI。
- `packages/ai-search/**` 与 Shell/Slides Genspark 调用点：CLI 删除、搜索 fallback 和 unsupported 能力。
- Electron manifests、builder、updater、README/SECURITY/NOTICE/GitHub metadata、品牌 assets：全量产品身份与发布配置。

## Task 1：npm scope 与代码命名空间迁移

### 文件

- 修改：根 `package.json`、`package-lock.json`
- 修改：`apps/docs/package.json`、`apps/pdf/package.json`、`apps/sheets/package.json`、`apps/shell/package.json`、`apps/slides/package.json`
- 修改：`packages/agent-core/package.json`、`packages/ai-provider/package.json`、`packages/ai-search/package.json`、`packages/docx-engine/package.json`、`packages/electron-utils/package.json`、`packages/file-parse/package.json`、`packages/i18n/package.json`、`packages/pptx-engine/package.json`、`packages/pptx-render/package.json`、`packages/project-store/package.json`、`packages/ui/package.json`
- 修改：所有被 `rg -l '@genoffice/' apps packages .github package.json` 命中的 TypeScript、配置和测试文件

### 接口与验收

- 所有 workspace 名称和内部依赖改为 `@wiswork/*`；目录结构与公开导出不变。
- 根包名改为 `wiswork`，脚本中的 workspace selector 同步改名。
- RED：先运行扫描，证明代码/manifest 中存在 `@genoffice/`。
- GREEN：`rg '@genoffice/' package.json package-lock.json apps packages .github` 无结果；`npm install --package-lock-only` 成功；`npm run typecheck` 成功；不产生业务行为 diff。
- 提交：`chore: migrate workspace scope to WisWork`

## Task 2：共享 Logto OAuth 与加密会话

### 新建文件

- `packages/auth/package.json`
- `packages/auth/tsconfig.json`
- `packages/auth/vitest.config.ts`
- `packages/auth/src/config.ts`
- `packages/auth/src/oauth.ts`
- `packages/auth/src/session.ts`
- `packages/auth/src/electron.ts`
- `packages/auth/src/index.ts`
- `packages/auth/tests/oauth.test.ts`
- `packages/auth/tests/session.test.ts`
- `packages/auth/tests/electron.test.ts`

### 修改文件

- 根 `package.json`、`package-lock.json`
- `apps/shell/package.json`、`apps/docs/package.json`、`apps/sheets/package.json`、`apps/slides/package.json`
- `apps/shell/src/main/index.ts`
- `apps/docs/src/main/docs-main.ts`
- `apps/sheets/src/main/sheets-main.ts`
- `apps/slides/src/main/slides-main.ts`、`apps/slides/src/main/ai-ipc.ts`
- `apps/shell/src/shared/home-api.ts`、`apps/shell/src/preload/index.ts`
- `apps/docs/src/shared/ipc.ts`、`apps/docs/src/preload/index.ts`
- `apps/sheets/src/shared/{ipc-channels,desktop-api}.ts`、`apps/sheets/src/preload/index.ts`
- `apps/slides/src/shared/ipc.ts`、`apps/slides/src/preload/index.ts`
- `apps/shell/tests/app-settings.test.ts`、`apps/shell/tests/home-counts.test.ts`

### 接口

- `createAuthorizationRequest(): { url, state }`
- `consumeCallback(url): Promise<AuthSession>`
- `getAccountStatus(): AccountStatus`
- `getAccessToken(): Promise<string | null>`
- `refresh(): Promise<AuthSession>`
- `logout(): Promise<void>`
- `AccountStatus = { loggedIn: boolean; email?: string; userId?: string }`
- 登录进度：`launched | success | error`，错误只暴露稳定 code。

### TDD 与验收

- RED：state 不匹配、重复 callback、过期事务、非法 scheme、safeStorage 不可用、并发 refresh 等测试先失败。
- GREEN：PKCE 为 S256；callback 只消费一次；refresh 单飞；401 最多重试一次；token 不出现在状态对象或错误文本。
- Electron 测试覆盖 macOS `open-url` 与 Windows/Linux second-instance URL 提取。
- 定点：`npm test -w @wiswork/auth`、Shell/Docs/Sheets/Slides account IPC 测试。
- 周边：相关四个 workspace typecheck 与 tests。
- 提交：`feat: add WisWork Logto authentication`

## Task 3：WisModel provider 与 Agent Loop 接线

### 修改文件

- `packages/ai-provider/src/types.ts`
- `packages/ai-provider/src/providers.ts`
- `packages/ai-provider/src/stream.ts`
- `packages/ai-provider/src/chat.ts`
- `packages/ai-provider/src/index.ts`
- `packages/ai-provider/tests/providers.test.ts`
- `packages/ai-provider/tests/stream.test.ts`
- `packages/ai-provider/tests/chat.test.ts`
- `packages/ai-provider/tests/images.test.ts`
- `packages/ai-provider/tests/watchdog.test.ts`
- `apps/docs/src/main/docs-main.ts`
- `apps/sheets/src/main/sheets-main.ts`
- `apps/slides/src/main/ai-ipc.ts`
- `apps/docs/src/renderer/ai/transport.ts`
- `apps/sheets/src/renderer/ai/transport.ts`
- `apps/slides/src/renderer/ai/transport.ts`
- `apps/pdf/src/renderer/ai/transport.ts`

### 接口

- `AiProviderId` 包含 `wiswork`，默认 provider 为 `wiswork`。
- 固定 base URL 为 `https://wismodel-proxy-dev.atominnolab.com/api/v1`，默认模型为 `deepseek/deepseek-v4-flash-0731`。
- 新增主进程配置解析器，从 `WISWORK_MODEL_API_KEY` 返回请求级 config；任何 renderer settings 中的 apiKey/baseUrl 对 `wiswork` 无效。
- 缺少登录返回 `auth_required`；缺少模型 Key 返回 `model_credentials_missing`。

### TDD 与验收

- RED：先增加 URL/header、默认模型、renderer 覆盖拒绝、SSE 文本/tool call、401、取消和非流式 fallback 测试。
- GREEN：请求命中固定 `/chat/completions`；Authorization 只来自注入 config；消息、图片和 tool results 的 payload 与现有 Agent Loop 一致。
- 测试不得使用真实 endpoint 或真实 Key，统一 mock fetch。
- 定点：`npm test -w @wiswork/ai-provider`。
- 周边：Docs、Sheets、Slides、PDF AI transport tests 与 typecheck。
- 提交：`feat: route AI through WisModel`

## Task 4：移除 Genspark CLI 并明确降级能力

### 删除文件

- `packages/ai-search/src/gsk.ts`
- `packages/ai-search/tests/gsk-login.test.ts`
- `packages/ai-search/tests/gsk.test.ts`
- `packages/ai-search/tests/manual-gsk-e2e.mts`

### 修改文件

- `packages/ai-search/src/index.ts`、`packages/ai-search/package.json`、`packages/ai-search/tests/search.test.ts`
- `apps/shell/src/main/index.ts`、`apps/shell/electron-builder.cjs`
- `apps/slides/src/main/ai-ipc.ts`、`apps/slides/src/main/slides-main.ts`、`apps/slides/src/main/i18n-main.ts`
- `apps/docs/src/main/docs-main.ts`
- Docs/Sheets/Slides 的 shared IPC、preload、AI Panel 与 `strings-ai.ts`
- `apps/shell/src/{shared/home-api.ts,preload/index.ts,renderer/src/Home.tsx,renderer/src/strings.ts}`
- `tools/gen-third-party-notices.mjs`、`NOTICE`、`package-lock.json`

### 接口与验收

- ai-search 只保留 Serper/DuckDuckGo web/image search；无 CLI 解析、登录或文件上传能力。
- 图片生成、媒体分析、云端 slide generation 和 PDF 云转换返回稳定 `unsupported_feature`，UI 不展示可执行入口。
- RED：新增测试断言 unsupported 返回和 Genspark fallback 不再执行。
- GREEN：`npm ls @genspark/cli` 不包含依赖；打包配置无 `gsk` resources；`rg '@genspark/cli|resolveGskEntry|gskLogin|gskApiKey' package.json package-lock.json apps packages tools` 无活动代码结果。
- 定点：ai-search、Shell、Slides、Docs、Sheets 测试。
- 提交：`refactor: remove Genspark runtime dependencies`

## Task 5：全量产品品牌、发布元数据与视觉资产

### 修改文件

- 根 `package.json`、`README.md`、`NOTICE`、`SECURITY.md`、`CODE_OF_CONDUCT.md`、`CONTRIBUTING.md`
- `.github/ISSUE_TEMPLATE/{bug_report,feature_request,config}.yml`、`.github/workflows/ci.yml`
- `apps/{docs,pdf,sheets,shell,slides}/package.json`
- `apps/{docs,pdf,sheets,slides}/electron.vite.config.ts`
- `apps/shell/electron-builder.cjs`
- `apps/{docs,shell}/src/main/updater.ts`
- `apps/shell/src/main/{index,tab-manager}.ts`
- 各 app 的 renderer `index.html`、入口、用户可见 i18n strings 与品牌 CSS
- `tools/gen-third-party-notices.mjs`、`tools/fidelity-compare.mjs`、`tools/verify-slide-copy-across-decks.mjs`

### 资产

- 替换 `apps/shell/build/icon.png`、`icon-mac.png`、`icon.icns`、`icon.ico`
- 替换 `apps/shell/src/renderer/src/assets/{app-icon.png,genoffice-logo.svg}`，后者重命名为 `wiswork-logo.svg`
- 替换 Docs/Sheets/Slides renderer 中的 app icon；保持 send/file-format 等非品牌图标不变。
- 临时 “W” 主图由 imagegen 生成，最终项目引用的位图必须复制进 worktree；平台尺寸和格式由确定性工具派生并检查。

### 行为与验收

- app IDs 使用 `com.atominnolab.wiswork` 与四个子 ID；产品和 artifact 名使用 WisWork。
- update 变量改为 `WISWORK_UPDATE_URL`；旧变量只允许在显式迁移兼容分支中读取一次并有测试。
- userData 从 `GenOffice`/`AI Office` 复制到 `WisWork`，不删除源目录。
- RED：品牌扫描测试先列出旧产品标识和 bundle ID。
- GREEN：用户可见/构建配置无旧品牌；允许列表只包含迁移、历史文件格式兼容和设计文档。
- 运行 strings tests、app settings/updater tests、`npm run notices`、`npm run licenses`、`npm run build:all`。
- 提交：`feat: rebrand product as WisWork`

## Task 6：跨模块安全验证与文档收尾

### 修改文件

- `README.md`：开发登录、深链注册、`WISWORK_MODEL_API_KEY` 配置与开发限制。
- `SECURITY.md`：OAuth、safeStorage、主进程模型凭证、deep-link 与 unsupported 能力边界。
- 必要时更新设计/计划文档中的实测接口差异，但不得弱化已批准安全约束。

### 验证

- fresh `npm install` 或 lockfile consistency check。
- `npm run format:check`
- `npm run lint`
- `npm run typecheck`
- `npm test`；若当前 LibreOffice 安装仍不能读取简单 CSV，单独运行并记录该外部失败，同时要求其余测试全绿。
- `npm run build:all`
- `npm run notices` 与 `npm run licenses`
- `git diff --check`
- 密钥扫描：服务 Key 格式、refresh fixed code、Bearer token 与认证响应不得出现在 tracked files 或 staged diff。
- 品牌扫描：旧 scope、旧 app IDs、活动 Genspark 代码与资源为零；迁移允许列表逐项人工检查。
- 独立 broad reviewer 审查完整 `2234f41..HEAD`，Critical/Important 全部处理并按需复审。
- 提交：`docs: document WisWork development setup`（仅在产生文档改动时）。

## 迁移、发布与回滚顺序

1. 先完成 namespace，使后续新增包直接使用 `@wiswork/*`。
2. 认证与 provider 分别通过单元测试后再切换 app 默认路径。
3. Genspark 删除在 WisWork 登录与模型链路绿色后执行，避免无可用 AI 路径。
4. 品牌资产与发布配置最后切换，确保所有引用指向已存在文件。
5. 不自动推送、发布或合并；完整验证后由用户选择本地合并、创建 PR 或保留分支。
6. 回滚时保留 worktree/分支与旧 userData；切回起点提交即可恢复原应用行为。
