# PPT Agent 发布兼容矩阵（工程证据）

依据[完整方案](./wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md) §10 O6。此表描述当前协议实现和自动化证据，不代替 Windows/Mac PowerPoint Desktop/Web 的发布验收。

| Taskpane / Relay / PC              | 当前行为                                                                       | 工程证据                                                                                                                                 | 发布状态                                           |
| ---------------------------------- | ------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------- |
| v2 / v2 / v2，协商 PPT 能力        | 按交集启用 `presentation.v1`、附件、图片和权利声明能力；未协商能力在发送前拒绝 | `apps/office-addin/tests/relay-session.test.ts`、`services/wiswork-relay/tests/relay.rs`、`apps/shell/tests/office-relay-client.test.ts` | 工程路径已覆盖；真实宿主待验收                     |
| v2 / v2 / v2，PC 只提供 `agent.v1` | 保留普通 Agent；PPT 专用请求在 Taskpane 侧拒绝                                 | `apps/office-addin/tests/relay-session.test.ts`                                                                                          | 工程路径已覆盖；真实旧 PC 构建待验收               |
| v1 / v2 / v2                       | Relay 仍接受 v1 基础配对；仅提供旧版 `agent.v1` 协议                           | `services/wiswork-relay/tests/relay.rs`、`apps/office-addin/tests/relay-session.test.ts`                                                 | 工程路径已覆盖；真实旧 Taskpane 构建待验收         |
| v2 / v2 / v1                       | v1 PC 不能完成 v2 配对，Taskpane 最终超时；不能作为受支持混合版本发布          | `services/wiswork-relay/src/lib.rs` 的配对版本校验                                                                                       | **发布阻塞：需协调 PC 升级或实现明确的不兼容提示** |
| v2 / v1 / 任意 PC                  | v1 Relay 不支持 v2 配对；没有静默降级                                          | v2 协议帧与能力协商要求                                                                                                                  | **发布阻塞：必须先升级 Relay**                     |

构建版本探测、Manifest、能力协商分别解决静态资源、协议能力和会话入口的问题；三者不能互相替代。发布时应先部署支持 v2 的 Relay 和 PC，再发布 v2 Taskpane，并保留旧 Taskpane 的回滚包。`version.json`、`taskpane.html` 与其哈希资源需原子发布；真实 Win/Mac/Web 混合版本与端到端冒烟仍待执行。
