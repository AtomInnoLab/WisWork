# WisWork LaTeX AI Native 编辑器实施计划

日期：2026-08-05
状态：可执行
目标分支：`feature/wiswork-latex`
批准设计：`docs/superpowers/specs/2026-08-05-wiswork-latex-ai-editor-design.md`

## 1. 目标与非目标

目标是在 WisWork 中交付一个本地单用户 LaTeX 工作区：支持新建/导入目录项目、多文件源码编辑、受控 Tectonic 编译、PDF/SyncTeX、编译诊断，以及“AI 只读分析 → 多文件 Diff → 用户确认 → 原子写入 → 自动编译 → 一键撤销”的完整闭环。

本计划不实现云同步、分享、多人协作、评论、Git UI、在线模板市场、远程编译、自定义命令、任意 TeX Live 工具链或 WisUsage 计费。

## 2. 架构摘要

新增 `@wiswork/latex-project` 管理目录边界、安全导入和写入事务，新增 `@wiswork/latex-compiler` 管理固定版本 Tectonic、资源包和隔离编译，再由独立 `@wiswork/latex` Electron 应用组合成编辑器。Shell 只持有项目目录身份和 WebContentsView 生命周期；文件访问、编译和确认写入都由 LaTeX 主进程执行，renderer 不持有任意文件系统或进程能力。

Tectonic sidecar 随平台安装包发布，完整 bundle 首次使用时从清单中的固定 HTTPS 地址下载并校验 SHA-256，之后编译强制使用本地 bundle、隔离项目副本、`--only-cached`、`--untrusted` 和 `--synctex`。

## 3. 全局约束

- 所有运行时代码和包使用 `@wiswork/*`、`WisWork`、`com.atominnolab.wiswork.*` 命名。
- 不提交、记录或渲染 WisModel 服务 Key、OAuth refresh fixed code、access token 或 refresh token。
- renderer 只能使用 typed IPC；不能传入根目录、可执行文件、bundle URL、环境变量或任意命令参数。
- 所有文件操作以主进程保存的规范化真实项目根为边界；拒绝绝对路径、路径穿越、NUL、链接逃逸和二进制 AI 写入。
- Tectonic 使用 `execFile`/`spawn` 参数数组，不经过 shell；编译在校验后的临时副本中执行。
- 每个行为变更先写失败测试，再写最小实现使其通过；每个任务完成后运行该任务的聚焦测试和 typecheck。
- 不改写用户未授权的已有修改；不删除项目源码；缓存与临时内容必须可重建。
- Tectonic 版本、平台资产 URL、sidecar SHA-256、bundle URL 和 bundle SHA-256 必须在进入打包任务前全部固定并由测试校验，禁止 `latest`。
- `feature/wiswork-latex` 是基于 `feature/wiswork-integration` 的叠加分支；前置 PR 合并后再变基到 `main`。

## 4. 关键接口

跨包只暴露以下稳定概念，具体实现保持内部化：

```ts
export interface LatexProjectHandle {
  id: string
  displayName: string
  rootPath: string // 仅主进程可见；不得越过 preload
  mainFile: string | null // POSIX 风格项目相对路径
}

export interface LatexCompileRequest {
  projectId: string
  revision: string
  mainFile: string
}

export interface LatexCompileResult {
  revision: string
  status: 'success' | 'error' | 'cancelled' | 'timeout'
  pdfPath?: string // 仅主进程内部使用
  synctexPath?: string
  diagnostics: LatexDiagnostic[]
  log: string
}

export interface LatexEditProposal {
  id: string
  projectId: string
  expiresAt: number
  files: Array<{
    path: string
    beforeSha256: string | null
    afterText: string
  }>
}
```

renderer 只接收项目 ID、展示名、相对路径、诊断和受控 PDF URL。提案确认绑定完整规范化内容、基线哈希、一次性 ID 和过期时间；任何变化都会使授权失效。

## 5. 文件责任图

### 新建

- `packages/latex-project/`：目录项目、路径策略、发现、导入、原子写入、快照和提案事务。
- `packages/latex-compiler/`：资产清单、bundle 安装、隔离工作区、编译队列、诊断和 SyncTeX。
- `packages/pdf-viewer/`：从现有 PDF 应用抽取的最小只读 PDF.js viewer、懒渲染和定位接口。
- `apps/latex/`：Electron main/preload、typed IPC、项目 session、CodeMirror 工作台、PDF 预览和 AI skill。
- `tools/tectonic/manifest.json`：固定 Tectonic 与 bundle 版本、平台 URL、SHA-256、许可证来源。
- `tools/fetch-tectonic.mjs`：CI/本地按清单下载和校验 sidecar，不接受调用方覆盖 URL。
- `e2e/latex-shell.spec.ts`、`e2e/latex-dirty-close.spec.ts`、`e2e/latex-tab-restore.spec.ts`：Shell 级用户流程。
- `apps/shell/src/main/tab-session.ts`：LaTeX 目录标签的原子持久化与恢复。

### 修改

- `package.json`、`package-lock.json`：workspace 脚本、CodeMirror/pdfjs/ZIP 依赖和锁文件。
- `apps/shell/package.json`、`apps/shell/electron-builder.cjs`：LaTeX 模块与 Tectonic sidecar 打包。
- `apps/shell/src/shared/tabs-api.ts`、`home-api.ts`：`latex` 标签和独立目录项目 API。
- `apps/shell/src/main/index.ts`、`tab-manager.ts`：runtime 配置、菜单、最近项目、打开/去重、dirty-close。
- `apps/shell/src/preload/index.ts`、`apps/shell/src/renderer/src/Home.tsx`、`TabBar.tsx`、`strings.ts`、`counts.ts`、`home.css`：Home 和标签 UI。
- `apps/shell/tests/tab-manager.test.ts`、`strings.test.ts`、`home-counts.test.ts`：Shell 契约回归。
- `tools/gen-third-party-notices.mjs`、`tools/check-licenses.mjs`、`tools/branding.test.mjs`：新模块、sidecar 许可证和品牌检查。
- `.github/workflows/package-macos.yml`：固定版 arm64 Tectonic 获取、架构校验和产物校验。
- `packages/project-store/src/types.ts`、`store.ts`、`index.ts`、`tests/store.test.ts`：增加目录资源到聊天的适配，不改变现有 Office 文件 API。

## 6. 交付任务

### 任务 1：目录项目内核与安全文本访问

交付：能够创建/打开 LaTeX 目录项目，识别主文件，并在项目边界内列出、读取和原子保存文本。

文件：

- 创建 `packages/latex-project/package.json`、`tsconfig.json`、`vitest.config.ts`。
- 创建 `packages/latex-project/src/types.ts`、`path-policy.ts`、`project.ts`、`main-file.ts`、`atomic-write.ts`、`index.ts`。
- 创建 `packages/latex-project/tests/path-policy.test.ts`、`main-file.test.ts`、`project.test.ts`。
- 修改根 `package.json` 的 `test`、`typecheck` 脚本并更新 `package-lock.json`。

最小实施顺序：

1. 写失败测试覆盖绝对路径、`..`、NUL、目录/文件 symlink、大小上限、非 UTF-8、缺失文件和规范化相对路径。
2. 实现 `ProjectPathPolicy`，所有访问都通过根目录真实路径与逐段 `lstat` 校验。
3. 写失败测试覆盖主文件优先级：保存配置、`Tectonic.toml` inputs、根 `main.tex`、唯一 `\\documentclass`、歧义。
4. 实现 `discoverMainFile()` 和基础 article 模板的原子创建。
5. 实现文本列表/读取/原子保存；保存使用同目录临时文件、`fsync` 和 rename，并返回内容 SHA-256。

验收与证据：

- 红：`npm run test -w @wiswork/latex-project` 首次因缺失实现失败。
- 绿：上述测试全部通过；恶意路径测试确认项目外文件未被读取或修改。
- `npm run typecheck -w @wiswork/latex-project` 通过。

提交：`feat(latex): add secure local project core`

### 任务 2：安全导入、快照与确认写入事务

交付：目录扫描、ZIP 安全导入、AI 多文件提案确认和整批撤销具备独立可测事务语义。

文件：

- 创建 `packages/latex-project/src/import.ts`、`limits.ts`、`snapshot.ts`、`proposal.ts`。
- 创建 `packages/latex-project/tests/import.test.ts`、`snapshot.test.ts`、`proposal.test.ts` 和恶意 ZIP fixtures。
- 修改 `packages/latex-project/src/index.ts`、`package.json`、根 `package-lock.json`。

最小实施顺序：

1. 测试 ZIP traversal、绝对路径、symlink/hardlink、重复覆盖、设备节点、条目数、目录深度、单文件/总大小和压缩比限制。
2. 使用现有 JSZip 生态解析元数据，逐条排他写入新临时目录，全部成功后原子移动；失败清理临时目录。
3. 测试快照配额、固定当前回滚点和二进制资源复制；快照只写 userData 缓存。
4. 测试 proposal 过期、一次性消费、基线哈希冲突、路径越界、二进制目标、部分写入故障和完整撤销。
5. 实现事务日志、临时文件批量准备、提交/回滚和一键恢复；MVP 拒绝删除操作。

验收与证据：

- 红：每组安全 fixture 在实现前导致测试失败。
- 绿：所有导入失败都不创建目标项目；所有 proposal 失败都保持整批文件原样。
- `npm run test -w @wiswork/latex-project && npm run typecheck -w @wiswork/latex-project` 通过。

提交：`feat(latex): add safe import and edit transactions`

### 任务 3：固定资产清单与首次 bundle 安装

交付：固定版本 Tectonic 资产和首次 bundle 下载具备可验证、可取消、失败不信任的安装状态机。

文件：

- 创建 `tools/tectonic/manifest.json`、`tools/fetch-tectonic.mjs`、`tools/tectonic/README.md`。
- 创建 `packages/latex-compiler/package.json`、`tsconfig.json`、`vitest.config.ts`。
- 创建 `packages/latex-compiler/src/manifest.ts`、`bundle-installer.ts`、`errors.ts`、`index.ts`。
- 创建 `packages/latex-compiler/tests/manifest.test.ts`、`bundle-installer.test.ts`。
- 修改根 `package.json`、`package-lock.json`、`.gitignore`。

最小实施顺序：

1. 从 Tectonic 官方 release 固定 `0.16.9` 的目标平台资产；从 Tectonic 官方配置固定兼容 bundle。把最终 URL、长度和 SHA-256 写入清单，不使用重定向后的可变 `latest` 地址。
2. 测试清单 schema、HTTPS host allowlist、每个平台唯一资产、64 位 SHA-256、许可证字段和 bundle 唯一性。
3. 测试安装状态 `missing/downloading/ready/error`、临时文件、校验失败、取消、重试、并发单飞和原子替换。
4. 实现下载器；网络仅用于清单固定 URL，日志只记录资产 ID、字节数和稳定错误码。
5. `tools/fetch-tectonic.mjs --platform darwin-arm64` 下载后再哈希并执行 `tectonic --version`；缓存目录不纳入 Git。

验收与证据：

- 红：篡改测试服务响应时安装器拒绝进入 `ready`。
- 绿：模拟下载测试全部通过；真实 arm64 资产在 macOS CI 中输出预期固定版本。
- `npm run test -w @wiswork/latex-compiler && npm run typecheck -w @wiswork/latex-compiler` 通过。

提交：`feat(latex): pin and verify tectonic assets`

### 任务 4：隔离 Tectonic 编译、诊断与 SyncTeX

交付：从磁盘 revision 创建隔离副本，安全运行单实例编译并返回结构化结果。

文件：

- 创建 `packages/latex-compiler/src/workspace.ts`、`runner.ts`、`queue.ts`、`diagnostics.ts`、`synctex.ts`。
- 创建 `packages/latex-compiler/tests/workspace.test.ts`、`runner.test.ts`、`queue.test.ts`、`diagnostics.test.ts`、`synctex.test.ts`。
- 创建 `packages/latex-compiler/tests/fixtures/` 中的最小成功、语法错误、多文件和 SyncTeX 项目。
- 修改 `packages/latex-compiler/src/index.ts`。

最小实施顺序：

1. 测试隔离复制拒绝链接、越界和超限文件，且编译工作目录不等于用户项目目录。
2. 测试固定参数数组包含 `--untrusted --only-cached --synctex --bundle <fixed> --outdir <isolated>`，并确认 shell 永不启用、环境为最小允许列表且设置 `TECTONIC_UNTRUSTED_MODE=1`。
3. 用假 sidecar 测试成功、非零退出、总超时、无输出超时、输出上限、取消和完整进程树终止。
4. 实现每项目一个活动任务；相同 revision 合并，新 revision 取消旧任务，迟到结果不得覆盖新 PDF。
5. 把日志解析为文件/行/列/严重级别诊断；解析 SyncTeX 正反向查询并把隔离路径映射回项目相对路径。
6. 仅复制 PDF、`.synctex.gz` 和裁剪日志到项目缓存；无论结果如何清理隔离工作区。

验收与证据：

- 红：假 sidecar 越权参数、挂起和大输出测试先失败。
- 绿：全部假进程测试通过；安装 Tectonic 的环境中最小 fixture 能以断网/only-cached 模式生成 PDF 和 SyncTeX。
- `npm run test -w @wiswork/latex-compiler && npm run typecheck -w @wiswork/latex-compiler` 通过。

提交：`feat(latex): add isolated tectonic compiler`

### 任务 5：LaTeX Electron 主进程、preload 与项目 session

交付：独立 `@wiswork/latex` 应用可以通过窄化 IPC 创建/打开项目、编辑文件和编译，且可以被 Shell 作为 WebContentsView 托管。

文件：

- 创建 `apps/latex/package.json`、`tsconfig.json`、`vitest.config.ts`、`electron.vite.config.ts`、`vite.renderer.config.ts`。
- 创建 `apps/latex/src/shared/ipc.ts`、`apps/latex/src/main/latex-main.ts`、`project-session.ts`、`ipc.ts`、`index.ts`、`apps/latex/src/preload/index.ts`。
- 创建 `apps/latex/tests/ipc.test.ts`、`project-session.test.ts`、`dirty-close.test.ts`。

产生接口：

- `configureLatexRuntime({ preloadPath, rendererUrl, rendererFile, tectonicPath, userDataPath })`
- `createLatexView(projectPath)`
- `latexQueryDirty(webContents)`、`requestLatexClose(webContents, parentWindow)`
- 项目打开/重命名事件，供 Shell 同步标题和 recent project。

最小实施顺序：

1. 写 sender 所属关系、payload schema、大小上限、项目 ID 到 session 映射和 renderer 不得提供根路径的失败测试。
2. 实现每个 WebContents 一个 project session；销毁时取消 watcher、编译和下载订阅。
3. 注册项目/文件/编译/SyncTeX/提案 typed IPC；PDF 使用受控 `wiswork-latex-pdf://<session>/<revision>` protocol，不开放任意 `file://`。
4. 实现磁盘 watcher、dirty 查询和关闭提示；外部冲突不自动覆盖编辑缓冲区。
5. standalone main 只用于开发；正式协议注册和全局登录仍由 Shell 独占。

验收与证据：

- 红：伪造 sender、项目 ID、路径和 bundle URL 的测试先失败。
- 绿：所有越权调用返回稳定错误码，项目外文件保持不变；clean/dirty/cancel/discard 路径通过。
- `npm run test -w @wiswork/latex && npm run typecheck -w @wiswork/latex` 通过。

提交：`feat(latex): add secure electron project host`

### 任务 6：CodeMirror 工作台、共享只读 PDF viewer 和编译反馈

交付：用户可在四区界面编辑多文件、自动保存、编译、查看 PDF/错误并进行 SyncTeX 定位。

文件：

- 创建 `apps/latex/src/renderer/index.html`、`main.tsx`、`App.tsx`、`styles.css`、`env.d.ts`。
- 创建 `apps/latex/src/renderer/editor/LatexEditor.tsx`、`editor-state.ts`、`latex-language.ts`。
- 创建 `apps/latex/src/renderer/project/ProjectTree.tsx`、`OpenTabs.tsx`。
- 创建 `apps/latex/src/renderer/compile/CompilePanel.tsx`、`diagnostics.ts`。
- 创建 `apps/latex/src/renderer/pdf/PdfPreview.tsx`、`synctex.ts`。
- 创建 `packages/pdf-viewer/package.json`、`tsconfig.json`、`src/ReadonlyPdfViewer.tsx`、`src/types.ts`、`src/index.ts`、`tests/viewer-state.test.ts`。
- 修改 `apps/pdf/src/renderer/App.tsx`，使其组合共享 viewer，而不是继续拥有私有页面渲染器。
- 创建 `apps/latex/src/renderer/i18n/strings.ts`、`locale.tsx`。
- 创建 `apps/latex/tests/editor-state.test.ts`、`diagnostics.test.ts`、`synctex-ui.test.ts`、`strings.test.ts`。
- 修改 `apps/latex/package.json`、根 `package-lock.json`。

最小实施顺序：

1. 安装 CodeMirror 6 核心、lint 和经许可证/维护状态审查的 LaTeX language extension；安装与现有 PDF 应用同版本的 `pdfjs-dist`。
2. 测试三层版本状态：编辑缓冲、磁盘 SHA、已编译 revision；过期编译结果不得刷新预览。
3. 实现项目树、多标签、行号、历史、搜索、括号匹配、LaTeX 高亮、诊断 gutter 和快捷键。
4. 实现去抖保存、显式编译、保存后自动编译、取消、日志面板和错误点击跳转。
5. 从现有 PDF App 抽取只读的文档 bytes、页面懒渲染、页码/缩放、点击坐标和定位 API 到 `@wiswork/pdf-viewer`；PDF App 与 LaTeX App 都组合该包，并沿用 worker/cmaps/fonts/wasm 静态资源策略。
6. 实现源码点击到 PDF 和 PDF 双击到源码；映射失败时退化为文件/行跳转。

验收与证据：

- 红：revision 竞争、诊断映射和 SyncTeX 坐标测试先失败。
- 绿：组件/状态测试通过，`npm run build -w @wiswork/latex` 生成 renderer 和 preload/main 产物。
- 手工 smoke：多文件编辑、编译错误修复、PDF 刷新和双向定位各完成一次。

提交：`feat(latex): build source and pdf workbench`

### 任务 7：Shell 标签、Home、最近项目与恢复

交付：LaTeX 成为 WisWork 第五种工作区，可从 Home/菜单新建或导入，目录标签可去重、关闭保护和重启恢复。

文件：

- 修改 `apps/shell/src/shared/tabs-api.ts`、`home-api.ts`。
- 修改 `apps/shell/src/main/tab-manager.ts`、`index.ts`、创建 `tab-session.ts`。
- 修改 `apps/shell/src/preload/index.ts`。
- 修改 `apps/shell/src/renderer/src/Home.tsx`、`TabBar.tsx`、`strings.ts`、`counts.ts`、`home.css`。
- 创建 `apps/shell/src/renderer/src/assets/file-tex.svg`、`apps/shell/src/main/assets/menu-tex.png`、`menu-tex@2x.png`。
- 修改 `apps/shell/package.json`、根 `package.json`、`package-lock.json`。
- 修改/创建 `apps/shell/tests/tab-manager.test.ts`、`tab-session.test.ts`、`strings.test.ts`、`home-counts.test.ts`。

最小实施顺序：

1. `TabKind` 加入 `latex`，测试目录 basename 标题、真实路径去重、激活、clean/dirty/cancel/discard 和关闭销毁。
2. runtime 加 `LATEX_OUT`（开发 `apps/latex/out`，打包 `resources/modules/latex`）及开发端口 `5177`。
3. Home API 增加独立 `LatexRecentProjectEntry`、`newLatexProject`、`importLatexProject`、`openLatexProject`、`latexRecents`；不复用单文件的 rename/duplicate/delete 行为。
4. 菜单、Home 快捷卡、导入入口、最近项目区和所有 locale 加入 LaTeX；测试 locale 键集合。
5. `tab-session.ts` 原子保存 `userData/open-tabs.json`；启动只恢复仍存在、校验通过的 LaTeX 目录，并恢复活动标签。
6. 根 `dev/test/typecheck/build:all` 和 Shell dependency 加入 `latex-project`、`latex-compiler`、`pdf-viewer`、`latex` 四个新 workspace。

验收与证据：

- 红：现有 tab/home tests 在新增期望下失败。
- 绿：Shell 聚焦测试、`npm run typecheck -w @wiswork/shell` 和 `npm run build:all` 通过。
- E2E：`latex/out` WebContentsView 出现，tab 标题正确，重启恢复同一目录且无重复 tab。

提交：`feat(shell): integrate latex project workspace`

### 任务 8：AI 只读工具、Diff 确认、编译验证与撤销

交付：现有 WisModel Agent 能分析 LaTeX 项目并通过一次性提案完成受控多文件修改。

文件：

- 创建 `apps/latex/src/renderer/ai/latex-skill.ts`、`tools.ts`、`transport.ts`、`AiPanel.tsx`、`ProposalReview.tsx`。
- 创建 `apps/latex/tests/ai-tools.test.ts`、`proposal-review.test.ts`、`ai-flow.test.ts`。
- 修改 `apps/latex/src/shared/ipc.ts`、`src/main/ipc.ts`、`src/renderer/App.tsx`。
- 修改 `packages/project-store/src/types.ts`、`store.ts`、`index.ts`、`tests/store.test.ts`。

最小实施顺序：

1. 定义自动允许工具：`list_project_files`、`search_project_text`、`read_project_text`、`get_compile_diagnostics`、`compile_project`；测试路径、文件数、文本长度和输出裁剪。
2. 定义 `propose_project_edits`，只创建规范化提案，不直接写文件；工具结果返回 proposal ID 和摘要，并明确返回 `mutated: false`。现有 `AgentLoop` 不提供等待确认的暂停点，因此模型绝不获得 apply/write 工具。
3. UI 展示逐文件 Diff，可整批取消或取消选择文件；任何内容变化都生成新 proposal ID。真正写入只能由用户点击触发独立的一次性确认 IPC，不能由 AgentLoop 继续执行。
4. 用户确认后主进程调用任务 2 的事务 API，写前快照、写后自动编译；编译错误只反馈，不自动修复。
5. 实现一键撤销整次提案并重新编译；测试撤销后的文件哈希和 PDF revision。
6. 为 project-store 增加目录 resource key/chat 适配器，复用消息和 timeline，不改变 Office 文件的现有映射。
7. 复用 `@wiswork/agent-core`、`createIpcTransport` 和现有 WisModel main IPC；确认服务 Key 仍只存在主进程环境。

验收与证据：

- 红：未经确认写入、重复消费、过期提案、外部并发修改和二进制目标测试先失败。
- 绿：AI 可读但不能直接写；确认后一次写入、一次自动编译；撤销恢复全部文件。
- `npm run test -w @wiswork/latex -w @wiswork/project-store` 与对应 typecheck 通过。

提交：`feat(latex): add confirmed ai project edits`

### 任务 9：打包、许可证与 macOS arm64 产物

交付：完整 LaTeX 模块和正确架构 Tectonic sidecar 进入 WisWork macOS 测试安装包，bundle 仍保持首次下载。

文件：

- 修改 `apps/shell/electron-builder.cjs`。
- 修改 `.github/workflows/package-macos.yml`。
- 修改 `tools/gen-third-party-notices.mjs`、`tools/check-licenses.mjs`、`tools/branding.test.mjs`。
- 修改 `apps/shell/build/THIRD-PARTY-NOTICES.txt`（由生成器产生）。
- 创建/更新 `docs/development/latex.md` 和相关 README 入口。

最小实施顺序：

1. builder 顶层资源加入 `apps/latex/out -> modules/latex`；macOS 原生资源加入 `native/tectonic`，不加入完整 bundle。
2. workflow 在 package 前按清单下载/校验 arm64 sidecar，执行 `file` 和 `tectonic --version`；package 后再次检查 app 内 sidecar 和 LaTeX renderer。
3. workflow paths 覆盖 builder、LaTeX app、compiler、清单和 lockfile；保留现有 xlsx sidecar 检查。
4. NOTICE/许可证检查加入 Tectonic、CodeMirror LaTeX language extension及新增依赖；品牌测试加入 `WisWork LaTeX` 和 app ID。
5. 文档说明首次 bundle 下载、缓存位置、离线模式、不支持功能、清理和故障诊断。

验收与证据：

- `npm run notices && npm run licenses && npm run test:branding` 通过。
- `npm run build:all` 后存在 `apps/latex/out/renderer/index.html`。
- macOS artifact 内 `native/tectonic` 为 arm64，`modules/latex/renderer/index.html` 存在，DMG/ZIP 校验和生成。
- 安装后首次编译显示 bundle 下载；完成后断网重启仍可编译基础项目。

提交：`build(macos): package latex and tectonic`

### 任务 10：E2E、全量回归与交付审查

交付：核心用户旅程、权限边界和现有 Office 能力均有通过证据，分支可以进入代码审查。

文件：

- 创建 `e2e/latex-shell.spec.ts`、`latex-dirty-close.spec.ts`、`latex-tab-restore.spec.ts`、`latex-ai-edit.spec.ts`。
- 修改 `e2e/home.spec.ts`、`e2e/helpers.ts`（仅增加稳定 LaTeX helper/data-testid 支持）。
- 按发现修复相关测试，不为通过测试放宽安全边界。

最小实施顺序：

1. E2E 覆盖新建、导入、编辑保存、编译错误、成功 PDF、dirty-close、重启恢复、AI Diff 确认和撤销。
2. 安全回归覆盖 renderer 越权 IPC、项目外路径、恶意 ZIP、提案重放和命令参数注入。
3. 运行格式、lint、全量 test/typecheck/build、E2E、notices/licenses/branding 和秘密扫描。
4. 检查 `git diff --check`、变更范围、提交边界和设计逐条覆盖；对高风险边界执行独立代码审查。

最终验证命令：

```bash
npm run format:check
npm run lint
npm test
npm run typecheck
npm run build:all
npm run test:e2e
npm run notices
npm run licenses
npm run test:branding
git diff --check
```

秘密扫描仅使用变量名和通用模式，不在命令或日志中复述真实凭证。最终手工验收至少覆盖 macOS arm64 的首次下载、离线复编译、恶意项目拒绝和 AI 撤销。

提交：`test(latex): verify end-to-end project workflow`

## 7. 迁移、回滚与发布

- 不迁移或改写现有 Docs/Sheets/Slides/PDF 文件；LaTeX recent projects 和 open tabs 使用新增 schema。
- 新增持久化文件必须带 schema version，并在未知新版本时只读失败，不覆盖。
- 功能回滚可移除 Shell 的 LaTeX 入口和打包资源；用户源码目录不受影响，userData 下 bundle、缓存和快照可安全遗留或由显式清理入口删除。
- 资产清单升级单独提交；先下载校验新 bundle，成功后切换指针，保留旧 bundle 直到没有活动编译。
- 首次发布仅提供 macOS arm64 测试包；Windows/Linux sidecar 和进程树终止验证完成前不宣称支持。
- 正式签名/公证流程保持 fail-closed；`WISWORK_UNSIGNED_MAC_BUILD=1` 仅用于现有临时测试 artifact。

## 8. 计划自检

- 设计中的项目、导入、编译、编辑/PDF、Shell、AI、打包和安全边界均对应独立任务。
- 没有任务依赖后置接口：项目内核先于编译，编译先于 UI，确认事务先于 AI 写工具，应用接口先于 Shell。
- 所有外部资产在任务 3 固定并校验，不把 URL/哈希留到发布时临时决定。
- 每个行为任务均包含失败测试、通过证据和独立提交。
- 回滚、schema、缓存、许可证、签名、跨平台声明和秘密扫描均有明确步骤。
