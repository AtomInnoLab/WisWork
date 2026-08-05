# WisWork 品牌、Logto 登录与 WisModel 接入设计

日期：2026-08-05
状态：已批准
目标分支：`feature/wiswork-integration`

## 1. 目标

将当前 GenOffice 仓库改造成 AtomInnoLab 自有产品 WisWork：完成全量品牌替换，以 Logto OIDC 取代 Genspark CLI 登录，并将本地 Agent Loop 的模型调用切换到 WisModel 开发代理。现阶段保持 Docs、Sheets、Slides、PDF 与统一 Shell 的本地编辑能力不变。

成功标准：

- 用户可以从 WisWork 打开系统浏览器登录，通过 `wiswork://oauth/callback` 返回桌面应用。
- 主进程可以验证 OAuth state、完成 code 交换、加密保存会话、刷新会话并退出。
- 四个编辑器的 Agent Loop 统一使用 WisModel OpenAI-compatible SSE，文本、图片输入、取消、超时和 tool-calling 行为保持兼容。
- 代码、构建产物配置、产品文案和图标不再以 GenOffice/Genspark 作为产品身份。
- 源码、设置文件、renderer、日志和 Git 历史中不出现服务级模型 Key、refresh fixed code、访问令牌或刷新令牌。

## 2. 非目标

- 本阶段不接 Gateway 模型转发，不接 WisUsage 用户级计费。
- 不实现新的图片生成、媒体分析或云端 PDF 转换服务。
- 不设计正式生产环境的服务 Key 分发方案；开发直连模式必须显式提供主进程环境变量。
- 不改写现有 Agent Loop、文档引擎、表格引擎或幻灯片引擎的业务语义。

## 3. 已选方案与备选方案

### 已选：共享认证模块 + 专用 WisWork provider

建立 `@wiswork/auth` 作为纯协议与状态模块，由 Electron 主进程注入浏览器打开、深链注册、加密存储和网络能力。`@wiswork/ai-provider` 新增固定的 `wiswork` provider，复用现有 OpenAI-compatible 转换器和 SSE 解析器。

该方案把登录会话、模型服务凭证和 renderer 隔离开，且未来可以将 provider 基址切换到 Gateway，而不需要改 Agent Loop。

### 未选：原地改写 Genspark 函数

直接将 `gskLogin`、`gskApiKey` 和 `genspark` provider 改名，初始改动较小，但会继续把认证、搜索 CLI 和模型凭证耦合在 `ai-search` 内，无法形成清晰的安全边界。

### 未选：保留多 provider 设置并使用 custom provider

现有 custom provider 能调用 WisModel，但会把模型 Key 暴露给 renderer 和 `ai-settings.json`，不满足桌面端密钥边界，也无法统一登录门禁。

## 4. 品牌与包标识

- 产品名：`WisWork`
- npm workspace scope：`@wiswork/*`
- Shell app ID：`com.atominnolab.wiswork`
- 子应用 app ID：
  - `com.atominnolab.wiswork.docs`
  - `com.atominnolab.wiswork.sheets`
  - `com.atominnolab.wiswork.slides`
  - `com.atominnolab.wiswork.pdf`
- OAuth 深链：`wiswork://oauth/callback`
- 包名、导入路径、构建配置、安装包名、更新环境变量、README、NOTICE、安全报告入口和用户可见文案统一使用 WisWork 命名。
- 现有目录名 `apps/docs`、`apps/sheets`、`apps/slides`、`apps/pdf` 保持不变，避免没有产品价值的路径迁移。
- 旧 `GenOffice` 与 `AI Office` userData 仅做一次复制式迁移；不删除旧数据。
- 临时视觉系统采用简洁 “W” 标志；生成位图后派生各平台所需 PNG/ICO/ICNS，已有代码原生 SVG 以确定性方式替换。

## 5. 认证架构

### 5.1 配置

开发环境默认值：

- Authorization endpoint：`https://auth.dev.wispaper.ai/oidc/auth`
- Gateway callback：`https://gateway.dev.wispaper.ai/api/v1/auth/user/callback`
- Refresh endpoint：`https://gateway.dev.wispaper.ai/api/v1/auth/user/refresh`
- Client ID：`y3xpwx3ytskxf66p0wztm`
- Desktop redirect URI：`wiswork://oauth/callback`

非敏感地址和 client ID 可以提供开发默认值，并允许通过 `WISWORK_OAUTH_*` 环境变量覆盖。refresh fixed code 属于服务端配置，不进入桌面客户端。

### 5.2 登录流程

1. 主进程生成高熵 state、PKCE verifier 和 S256 challenge，并把待处理事务写入内存；事务有明确过期时间且只能消费一次。
2. 主进程构造 Logto authorization URL，scope 为 `openid profile email offline_access`，通过受控 `shell.openExternal` 打开。
3. Shell 注册 `wiswork` 协议。macOS 使用 `open-url`，Windows/Linux 使用 single-instance argv；两条路径进入同一解析器。
4. 解析器只接受精确 scheme、host/path、允许字段和最大长度，恒定时间比较 state，拒绝过期、重复或未知回调。
5. 主进程将 code、redirect URI 和 PKCE verifier 交给 Gateway callback，接收 WisPaper `CallbackParams`，至少包含 `token`、`refresh_token` 和 `user_id`。
6. 会话通过 Electron `safeStorage` 加密后写入应用 userData。`safeStorage` 不可用时登录失败，不降级为明文保存。
7. 登录状态 IPC 只向 renderer 返回 `loggedIn`、用户展示字段和错误类别，不返回任何 token。

### 5.3 刷新与退出

- 请求前检查 token 到期时间；临近到期时以单飞锁调用 refresh endpoint，POST JSON 至少包含 `refresh_token`。
- 刷新成功后原子替换加密会话；并发请求共享同一次刷新。
- 401 时最多执行一次刷新后重试，避免无限循环。
- `invalid_grant` 或无法解密时清除本地会话并返回 `auth_required`。
- 退出会清除加密会话、待处理 OAuth 事务和内存 token；不删除文档或项目数据。

## 6. 模型调用架构

### 6.1 Provider

- Provider ID：`wiswork`
- API base URL：`https://wismodel-proxy-dev.atominnolab.com/api/v1`
- 默认模型：`deepseek/deepseek-v4-flash-0731`
- 完整路径：`/chat/completions`
- 协议：OpenAI-compatible JSON 与 SSE
- 主进程请求头：`Authorization: Bearer <WISWORK_MODEL_API_KEY>`、`Content-Type: application/json`

服务 Key 只从主进程环境变量 `WISWORK_MODEL_API_KEY` 读取。缺失时返回可识别的 `model_credentials_missing` 错误；不读取 renderer 设置，不写磁盘，不写日志，也不使用用户 OAuth token 替代。

### 6.2 Agent Loop 兼容

复用现有 OpenAI-compatible 适配：

- system、user、assistant 与 tool results 的消息映射保持不变。
- 图片继续使用 data URL content part。
- tools 转换为 OpenAI function tools。
- SSE 累积 tool-call argument fragments，并在 finish 时转换为 `AgentToolCall`。
- `[DONE]`、`finish_reason`、取消、连接超时、流空闲超时和非流式 JSON fallback 保持现有行为。
- 错误响应只记录 provider、model、HTTP 状态与经过裁剪的非敏感错误文本。

登录是产品门禁，模型 Key 是开发服务凭证，两者职责分离；登录成功不代表模型凭证存在。

## 7. Genspark 能力处置

- 删除 Genspark 作为默认认证和模型 provider 的路径。
- 删除 `@genspark/cli` 运行时依赖、打包资源与登录配置读取。
- web/image search 保留 Serper 与 DuckDuckGo 可用链路。
- 依赖 Genspark CLI 的图片生成、媒体分析、云端幻灯片生成和 PDF 云转换暂时关闭，并返回稳定、可本地化的“不支持”错误。
- 不静默回退到用户自带模型 Key或其他外部 provider。

## 8. IPC 与信任边界

- renderer 只能调用 typed IPC：登录、退出、查询状态、开始/取消 AI 请求。
- 主进程验证 IPC payload、sender 和 request ID；renderer 不能传入 authorization URL、callback URL、模型 base URL 或请求头。
- OAuth callback 与普通外部链接使用不同的允许列表；现有 http/https 外链限制不放宽。
- 敏感值在错误、遥测、测试快照与 console 输出前统一脱敏。
- 模型请求仍由 Electron 主进程发出，避免 renderer CORS 与凭证暴露。

## 9. 失败处理

- 用户取消或 OAuth 超时：结束当前事务，保留已存在的有效会话。
- state/协议不匹配：拒绝回调并记录不含参数值的安全事件。
- callback/refresh 网络错误：返回可重试错误，不清除仍有效的旧 token。
- 会话过期且刷新失败：进入 logged-out 状态并引导重新登录。
- 模型 401/403：返回凭证错误，不回显服务 Key。
- 模型 429/5xx/超时：沿用流式错误与取消语义，Agent Loop 不执行不完整 tool call。
- unsupported Genspark feature：明确告知当前 WisWork 开发版本未配置该能力。

## 10. 测试与验证

- `@wiswork/auth`：URL 构造、state/PKCE、深链解析、重复回调、过期、刷新单飞、401 单次重试、加密存储失败和退出。
- `@wiswork/ai-provider`：固定基址、主进程 Key、SSE 文本、tool calls、图片消息、非流式 fallback、401、超时和取消。
- Electron IPC：token 不越过 preload，renderer 不能覆盖 URL/header，Shell 与 standalone app 行为一致。
- 品牌：扫描旧 scope、产品名、bundle ID、安装包名、更新变量和 Genspark 依赖；保留的历史迁移字符串必须有注释和允许列表。
- 供应链：`npm install`、license/notices 生成与检查。
- 仓库：完整 unit tests、typecheck、build:all、format check 与 diff check。
- 安全：扫描提交内容，确认不存在用户提供的模型 Key、refresh fixed code、token 或 refresh token。
- 手工验收：macOS 与 Windows 各验证一次冷启动登录、深链回调、重启恢复、刷新、退出、模型流和 tool call。

当前执行环境的 LibreOffice 24.2.7.2 无法加载最简单 CSV 或空白 XLSX，因此 pivot LibreOffice round-trip 属于外部环境未验证项；其他 Sheets 测试必须全绿。

## 11. 发布与回滚

- 所有工作位于 `feature/wiswork-integration` 独立 worktree；未经用户选择不合并 `main`、不推送、不发布。
- 开发构建必须显式设置 `WISWORK_MODEL_API_KEY`；生产打包流程在正式凭证方案完成前不得包含该变量或其值。
- OAuth 深链必须先在 Logto 注册，并确保操作系统安装包正确注册协议，再进行真实登录验收。
- 回滚通过停止分发 WisWork 构建并回到改造前提交完成；旧 GenOffice userData 未被删除，可继续由旧版本读取。
- 新 WisWork userData 写入失败时不得回写或破坏旧目录。

## 12. 风险与责任

- OAuth callback 参数或 `CallbackParams` 实际结构与本设计不一致时，由认证接口负责人确认契约后调整适配器；不得通过宽松 `any` 或忽略 state 继续。
- WisModel 对 OpenAI tool-calling 的兼容性由集成测试验证；不兼容时该能力应显式阻断，而不是伪装成文本成功。
- 开发直连服务 Key 的泄露风险由开发部署环境承担；客户端代码只负责不持久化、不渲染和不记录。
- 正式计费、限流与服务 Key 隐藏由后续 Gateway/WisUsage 集成负责。
