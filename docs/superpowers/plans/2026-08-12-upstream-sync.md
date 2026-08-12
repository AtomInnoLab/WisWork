# WisWork 上游同步执行计划

日期：2026-08-12
状态：执行中
目标分支：`codex/sync-upstream-2026-08-12`
基线：`AtomInnoLab/WisWork@51b892823706167e9dc9ce5067cd9e8b87529c27`
上游：`genspark-ai/genoffice@dc4d7e5927864498913b7ba42d0da06cc7cf628e`

## 目标与非目标

目标是将指定的 GenOffice 上游快照合入 WisWork，保留真实 merge ancestry，使后续同步只处理增量变化；吸收上游的文件格式、稳定性、跨平台和通用 UI 改进，同时维持 WisWork 的品牌、包作用域、Logto/WisPaper 认证、WisModel/WisUsage、LaTeX 应用和安全边界。

本次不发布、不直接合入 `main`、不自动推送、不执行 `npm audit fix`，也不恢复 Genspark CLI、Genspark 登录、Genspark 模型路由或 GenOffice 用户可见品牌。依赖审计在合并前已有 49 个 high severity 告警，作为既有风险记录，不在同步中扩大处理范围。

## 架构与全局约束

上游在文档格式引擎、解析渲染、通用稳定性修复上优先；WisWork 在身份、品牌、认证、模型、计费、权限和 LaTeX 产品能力上优先。冲突按能力所有权解决，不对整个文件机械选择 ours/theirs。新增 Markdown 应用可以吸收，但必须进入 WisWork 的包作用域、Shell 注册、AI transport 和品牌边界。

全局约束：

- 保留一个包含 `51b8928` 与 `dc4d7e5` 两个 parent 的 merge commit；禁止 squash 或重写共同历史。
- 所有 workspace 包和内部 import 使用 `@wiswork/*`；用户可见身份使用 WisWork。
- 不重新引入 `@genspark/cli`、Genspark 登录配置、Genspark 模型凭证或 renderer 可见服务密钥。
- 保留 `apps/latex`、`packages/auth`、`packages/project-store`、`packages/latex-project`、`packages/latex-compiler` 及其 Shell 集成。
- 上游新增 Markdown 应用不得挤占或替换 LaTeX 应用。
- package-lock 必须由最终 package manifests 确定性重建，不手工拼接冲突块。
- 本次主要是集成与行为保持，不适用为每个上游行为重新制造 RED；以合并前测试作为 characterization baseline，对新适配边界补充或更新测试，并执行确定性 GREEN 验证。

## Task 1：建立合并树与冲突清单

涉及：Git 历史、全部冲突文件。

执行：

1. 在当前同步分支执行 `git merge --no-ff --no-commit upstream/main`。
2. 保存未合并文件清单，并按 Docs、Sheets/PDF、Slides、Shell/共享基础设施分类。
3. 检查自动合并结果中是否出现 Genspark、`@genoffice`、应用注册丢失或 WisWork 安全边界回退。

验收：merge 处于可恢复状态；冲突集合已知；两个 parent SHA 与计划一致。

## Task 2：Docs 冲突与上游引擎能力

主要文件：`apps/docs/**`、对应 Docs 测试、`packages/docx-engine/**`。

原则：以上游 Docs 文件保真、分页、字体、形状和编辑修复为基础，重新保留 WisWork AI dock、主题、认证路由和 `@wiswork` import。上游删除的旧组件仅在 WisWork 仍有真实引用时保留。

验收：Docs 无冲突标记；`@wiswork/docs` 单测、typecheck 通过；WisWork AI 与主题测试继续存在并通过；上游新增 Docs 测试通过。

## Task 3：Sheets 与 PDF 冲突

主要文件：`apps/sheets/**`、`apps/pdf/**`、对应共享 IPC 和测试。

原则：吸收上游 xlsx sidecar、PDF 编辑和保真修复；保留 WisWork AI transport、右侧 dock、主题、包作用域和 Shell 协议。Rust sidecar 测试使用仓库内 `CARGO_TARGET_DIR`，避免共享 target 路径造成假失败。

验收：两个应用无冲突标记；除已记录的 LibreOffice 外部环境用例外，Sheets 测试通过；PDF 测试与 typecheck 通过。

## Task 4：Slides 冲突

主要文件：`apps/slides/**`、`packages/pptx-engine/**`、`packages/pptx-render/**`。

原则：吸收上游字体、布局、演示和格式保真修复；保留 WisWork AI transport、主题、包作用域和附件安全边界。

验收：Slides 无冲突标记； Slides、pptx-engine、pptx-render 测试与 typecheck 通过。

## Task 5：Shell、Markdown 与共享基础设施

主要文件：`apps/shell/**`、`apps/markdown/**`、`packages/agent-core/**`、`packages/ai-provider/**`、`packages/ai-search/**`、`packages/i18n/**`、`packages/ui/**`、根 `package.json`、lockfile、构建与 notice 工具。

原则：Shell 最终同时注册 Docs、Sheets、Slides、PDF、Markdown、LaTeX；保留 WisWork 登录、更新源、品牌、主题和 WisUsage；Markdown 通过 WisWork provider/transport，不依赖 Genspark CLI；共享 provider 吸收上游网络与流处理修复，但 WisWork 鉴权和错误语义优先。删除/修改冲突中不恢复已经移除的 Genspark 文件。

验收：workspace 和 lockfile 一致；branding test 不发现禁止标识；认证、AI provider、Shell、Markdown、LaTeX 测试和 typecheck 通过；构建配置包含六个应用且不丢失 Tectonic/LaTeX 资源。

## Task 6：集成审查、验证与 merge commit

1. 对每个模块单元进行独立审查，修复 Critical/Important 问题。
2. 扫描冲突标记、`@genoffice`、Genspark 运行时路径、旧品牌和敏感配置回退。
3. 运行 format check、lint、typecheck、完整单测、branding、licenses、notices 和 `build:all`；LibreOffice 已知项单独记录。
4. 检查最终 diff、workspace 清单、Shell 应用注册和 merge parent。
5. 创建一个保留上游 parent 的 merge commit，不推送、不合入 `main`，等待用户选择。

验收：没有未解决冲突或非预期生成物；适用验证通过；独立最终审查无 Critical/Important 问题；merge commit 的两个 parent 分别来自 WisWork 基线链和指定上游快照。

## 回滚、迁移、安全与发布

- 回滚：分支未合入 `main` 前直接保留或放弃该隔离分支；不修改用户数据或远端状态。
- 迁移：不得删除或改写现有 WisWork userData；新增 Markdown 设置与现有应用设置隔离。
- 安全：认证 token、模型服务凭证和 WisUsage 会话继续只在主进程受控边界中处理；上游网络代码必须经过现有 provider/IPC 契约。
- 发布：本次不发布。只有分支完整验证后，用户才能选择推送 PR 或保留分支。
