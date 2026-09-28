# PC 主动撤销 Relay 会话实施计划

依据：[设计](../specs/2026-09-28-pc-explicit-relay-revocation-design.md)。目标是主动关闭立即撤销 v2 会话；意外关闭仍可重附着。复用客户端已有关闭帧，不改网络消息格式、存储或认证策略。

## 任务 1：Relay 精确识别主动撤销

- 文件：`services/wiswork-relay/src/lib.rs`、`services/wiswork-relay/tests/relay.rs`。
- 先添加失败集成测试：已认证 PC 的 `1000/session_revoked` 关闭使同会话 Office 收到 `session_revoked`，原能力不能 `pc.resume`；普通关闭仍收到 `office.pc_offline` 且可恢复；其他会话仍可请求。
- 在 WebSocket 处理循环仅对已认证 PC 的精确关闭帧记录主动撤销，并把该布尔值传给清理函数。主动撤销时使用现有会话删除和通知逻辑；非主动分支不变。
- 验证：目标集成测试先红后绿，`cargo fmt --check`、`cargo test`；提交仅包含该任务文件。

## 任务 2：端到端回归与阶段记录

- 文件：`apps/shell/tests/office-relay-client.test.ts`、`docs/product/wiswork-ppt-agent-overall-progress-2026-09-23.md`（仅在现有测试不足时改 PC 测试）。
- 核对 PC 客户端 `revoke` 的精确关闭参数与池的多连接调用；验证定向和全量 PC 测试，并执行 Relay 完整套件。
- 记录 O0 工程闭环、真实宿主缺口和整体进度。完成独立代码审查，再提交记录。

## 发布与回滚

Relay 可先发布，旧 PC 主动关闭若未使用精确 reason 则沿原有短时恢复路径。新 PC 已使用该 reason。回滚 Relay 变更即可恢复旧行为。无迁移。
