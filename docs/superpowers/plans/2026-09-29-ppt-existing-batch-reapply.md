# 现稿批量 ChangeSet 重新应用实施计划

目标：按产品方案 v1.3 §15 阶段 5，让已撤销且备份完整的现稿批量变更可确认式重做，并继续可恢复、可撤销。非目标：跨类型重做、严格原子写入、替代真实 PowerPoint 验收。

架构：复用 `presentation-existing-batch-editing.ts` 的逐步写入与回读循环；`presentation-existing-batch.ts` 只增加有界的 `undone → applying` 转换。Taskpane 通过既有 `presentation-changes` 控制器路由新增动作，并在 `host-runtime.ts` 将该操作纳入受影响页 QA 失效范围。

全局约束：备份已释放或不可读时写前拒绝；非目标字段及目标当前值重校验；历史视觉审阅清除；所有宿主写入走提案确认；未知状态不得自动重放。当前工作区为隔离分支 `codex/ppt-agent-implementation`，基线 `1f76c771`。

## 任务 1：状态机与工具执行

文件：`apps/office-addin/src/skills/powerpoint/presentation-existing-batch.ts`、`presentation-existing-batch-editing.ts`，对应 `presentation-existing-batch.test.ts`、`presentation-existing-editing.test.ts`。

验收：已撤销记录在确认后逐步重做，重开后可再次撤销；释放/缺失备份、手工第三值与保留字段变化在写前拒绝；首步写入后回执丢失可恢复；转换不允许改操作、范围或备份身份。先运行新增测试并看到状态无效或工具不存在的 RED，再实现并得到 GREEN；提交一个有界功能提交。

## 任务 2：工作台入口和 QA 范围

文件：`apps/office-addin/src/agent/presentation-changes.ts`、`presentation-changes-card.tsx`、`host-runtime.ts`，对应控制器、卡片及宿主 QA 测试。

验收：`undone` 且备份未释放显示“重新应用”，按钮按 change ID 调用新工具；已释放备份不显示；提案确认后仅受影响页 QA 失效，历史审阅不可冒充当前 QA。先运行入口/QA 回归 RED，再实现 GREEN；提交一个有界功能提交。

## 任务 3：整体验证与记录

文件：`docs/product/wiswork-ppt-agent-overall-progress-2026-09-23.md`、必要的验收说明。

运行 Office 插件全套测试、类型检查、定向 lint、Prettier、`git diff --check`；独立复核任务 1/2 完整差异，修复重要问题后再复核。记录真实宿主尚未执行、工程百分比和下一步；不得以模拟测试宣称 100%。
