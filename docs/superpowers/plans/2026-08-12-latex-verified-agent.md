# LaTeX Verified Agent 实施计划

## Goal

交付一个从编辑上下文发起、过程可观察、提案可隔离验证、正式写入需确认且可撤销的 LaTeX AI 垂直闭环。

非目标：自动删除、任意命令、无人确认写入、长期任务存储、PDF 视觉 diff、多轮无人值守自修复。

架构上，renderer 只采集和展示上下文与证据；`ProjectSession` 持有提案并授权验证；编译包在其现有安全临时工作区内应用有界文本 overlay。验证不消费提案、不发布正式预览、不接触项目源码。

全局约束：保持现有项目路径策略、AI 敏感文件限制、一次性 proposal 授权、基线哈希、事务锁、快照、Tectonic 固定参数和 shell-escape 禁止策略。

## Task 1: 上下文与任务时间线

文件：

- 修改 `apps/latex/src/renderer/editor/LatexEditor.tsx`：报告有界选区与光标行。
- 新增 `apps/latex/src/renderer/ai/agent-context.ts`：上下文模型、边界和 prompt 序列化。
- 新增 `apps/latex/src/renderer/ai/task-timeline.ts`：工具调用到时间线状态的纯函数模型。
- 修改 `apps/latex/src/renderer/App.tsx` 与 `ai/AiPanel.tsx`：传递、移除、发送上下文并展示时间线。
- 修改 `CompilePanel.tsx`：允许把诊断发送给 AI。
- 新增/修改 `apps/latex/tests/agent-context.test.ts`、`task-timeline.test.ts`、相关 UI 测试。

验收：选区和诊断上下文有界、可见、可移除；prompt 明确区分用户指令与不可信上下文；工具开始/成功/失败/取消均可观察。

TDD：先写纯函数和组件契约测试并观察缺少模块/行为的 RED，再实现；运行新增测试、LaTeX 全套测试和 typecheck。

提交：`feat(latex): add contextual agent task timeline`

## Task 2: 编译工作区 overlay

文件：

- 修改 `packages/latex-compiler/src/workspace.ts`：定义并应用有界 text overlay。
- 修改 `packages/latex-compiler/src/runner.ts` 与 `index.ts`：把 overlay 传入隔离编译流程。
- 修改 `packages/latex-compiler/tests/workspace.test.ts`、`runner.test.ts`、hardening tests。

验收：overlay 可修改现有文本或创建允许的新文本文件；原项目不变化；路径穿越、绝对路径、符号链接目标、二进制/NUL、单文件/总量超限均在 spawn 前拒绝。

TDD：先加入行为和攻击面测试并确认 RED；最小实现后运行 latex-compiler 全套测试与 typecheck。

提交：`feat(latex-compiler): support bounded compile overlays`

## Task 3: 提案隔离验证 IPC

文件：

- 修改 `apps/latex/src/shared/ipc.ts`、`preload/index.ts`、`main/ipc.ts`：增加 proposal verify typed IPC。
- 修改 `apps/latex/src/main/project-session.ts`：验证 session 所有的未过期提案，调用 overlay compile，但不消费提案或发布正式预览。
- 修改 `apps/latex/tests/ipc.test.ts`、`project-session.test.ts`、`ai-flow.test.ts`。

验收：renderer 只能以当前项目和 proposal ID 请求验证；不能提交 overlay 文本、编译器参数或路径；验证不修改项目、不消费提案，验证后仍可正式应用一次；脏缓冲区和变化基线被拒绝。

TDD：先写 IPC schema、权限和 session 行为测试确认 RED；实现后运行 LaTeX 和相关内核测试与 typecheck。

提交：`feat(latex): verify AI proposals in isolation`

## Task 4: 验证式提案审阅

文件：

- 新增 `apps/latex/src/renderer/ai/diff.ts`：有界行级变更块模型。
- 修改 `ProposalReview.tsx`、`proposal-review.ts` 和 `AiPanel.tsx`：提案加载后触发验证，显示每文件变更块、理由占位与验证证据。
- 修改 `styles.css` 和相关测试。

验收：审阅显示文件级摘要和行级变更块；验证中/通过/失败状态清楚；验证失败仍可审阅但确认按钮采用高风险提示；应用后显示正式编译结果和与隔离验证的差异。

TDD：先写 diff 模型、状态机和渲染契约测试确认 RED；实现后运行 LaTeX 全套测试、typecheck 和 build。

提交：`feat(latex): review verified AI changes`

## Task 5: 整体验证与发布准备

- 更新 `docs/development/latex.md`，说明 proposal verification 的边界与排障。
- 运行 formatter/lint（若仓库脚本提供）、LaTeX/latex-project/latex-compiler 全套测试、相关 typecheck、LaTeX build。
- 运行仓库声明的完整 `npm test` 和 `npm run typecheck`；记录任何与本分支无关的环境或既有失败。
- 审查最终 diff，确认没有凭证、任意路径、任意命令或未经确认写入通道。
- 安排独立代码审查并修复 Critical/Important 发现。

验收：设计中的安全、回滚和验证标准均有测试证据；分支可由用户选择合并、创建 PR 或保留。

提交：`docs(latex): document verified agent workflow`（仅在文档未与前述提交同交付时）。

