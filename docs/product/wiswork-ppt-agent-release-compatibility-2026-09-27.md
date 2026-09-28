# PPT Agent 发布兼容矩阵（工程证据）

依据[完整方案](./wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md) §10 O6。此表描述当前协议实现和自动化证据，不代替 Windows/Mac PowerPoint Desktop/Web 的发布验收。

| Taskpane / Relay / PC              | 当前行为                                                                                                                                                                                 | 工程证据                                                                                                                                 | 发布状态                                                |
| ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------- |
| v2 / v2 / v2，协商 PPT 能力        | 按交集启用 `presentation.v1`、附件、图片和权利声明能力；未协商能力在发送前拒绝                                                                                                           | `apps/office-addin/tests/relay-session.test.ts`、`services/wiswork-relay/tests/relay.rs`、`apps/shell/tests/office-relay-client.test.ts` | 工程路径已覆盖；真实宿主待验收                          |
| v2 / v2 / v2，PC 只提供 `agent.v1` | 保留普通 Agent；PPT 专用请求在 Taskpane 侧拒绝                                                                                                                                           | `apps/office-addin/tests/relay-session.test.ts`                                                                                          | 工程路径已覆盖；真实旧 PC 构建待验收                    |
| v1 / v2 / v2                       | Relay 仍接受 v1 基础配对；仅提供旧版 `agent.v1` 协议                                                                                                                                     | `services/wiswork-relay/tests/relay.rs`、`apps/office-addin/tests/relay-session.test.ts`                                                 | 工程路径已覆盖；真实旧 Taskpane 构建待验收              |
| v2 / v2 / v1                       | Relay 向 v1 PC 返回版本不匹配错误并保持 Taskpane 配对待完成；已发布的旧 PC 会显示通用协议错误，Taskpane 最终超时。新版 PC 的错误文案仅覆盖直接 v1 领取的兼容测试路径，生产版使用 v2 协商 | `services/wiswork-relay/tests/relay.rs`、`apps/shell/tests/office-relay-client.test.ts`                                                  | **发布阻塞：必须先升级 PC；旧二进制不能原地获得新提示** |
| v2 / v1 / 任意 PC                  | v1 Relay 不支持 v2 配对；收到精确旧版协议错误时提示升级，不会静默降级                                                                                                                    | `apps/office-addin/tests/relay-session.test.ts`                                                                                          | **发布阻塞：必须先升级 Relay**                          |

构建版本探测、Manifest、能力协商分别解决静态资源、协议能力和会话入口的问题；三者不能互相替代。发布时应先部署支持 v2 的 Relay 和 PC，再发布 v2 Taskpane，并保留旧 Taskpane 的回滚包。`version.json`、`taskpane.html` 与其哈希资源需原子发布；真实 Win/Mac/Web 混合版本与端到端冒烟仍待执行。

灰度构建可设置 `VITE_WISWORK_PRESENTATION_ROLLOUT_PERCENT=0..100`（默认 100）。Taskpane 按已保存文稿的 Office 地址只读分桶；部分灰度中未保存或未命中的 PowerPoint 文稿在启动 Agent 前显示不可用状态。比例随新构建发布，回滚需要保留并切回旧构建及 Manifest。该开关只控制 PPT Agent 入口。可另设 `VITE_WISWORK_OFFICE_DIAGNOSTIC_SAMPLE_PERCENT=0..100`（默认 100），按运行 trace 对脱敏失败诊断远程抽样，本地诊断仍完整；抽样不是健康指标验收，也不能代表真实宿主兼容或部署后全链路冒烟已完成。

发布预检命令（部署前先用构建产物执行，部署后再加 `--deployed 1`）：

```bash
node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist
node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist --deployed 1 --rollout-percent 25 --diagnostic-sample-percent 10
PPT_AGENT_RELEASE_PC_TOKEN='<release-test-PC-token>' node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist --deployed 1 --rollout-percent 25 --diagnostic-sample-percent 10 --pairing 1
```

部署后预检要求显式填写本次发布的灰度比例和诊断采样比例，并与 `version.json` 中的构建元数据逐项核对；示例的 25% 和 10% 需替换为实际发布值。预检核对 Manifest 各 Office 资源地址及其本地文件、HTML 引用的样式文件、`version.json` 与编译入口的一致性、JS/CSS/worker 等运行时资源的哈希式文件名、除官方 Office.js 外无额外脚本、无源码映射、Relay `/office-relay/health` 的精确响应；本地构建的全部发布文件（含 CSS、worker、图标和 Manifest）形成大小与 SHA-256 清单。部署后逐一限量读取线上文件并比对完整字节，缺失、超量或篡改任一文件均失败。任何失败均以非零退出码阻断继续发布。CI 会构建真实 Taskpane 并执行产物预检，以及使用 LibreOffice 检查 PPTX 包回读。此命令不替代 v1/v2 配对协议测试（CI 的 Relay/Office 测试）及真实 PowerPoint 的附件、图片、页面写入、截图、恢复冒烟。诊断协议新增 `run`/`passed` 与 `agent_run_completed`；须先部署支持该事件的 Relay，再发布新版 Taskpane。事件只证明 Agent 流程结束，不代表 PPT 交付完成；真正完成率仍需 20 项任务与宿主证据。先保留旧版回滚包；发布顺序仍为 Relay、PC、Taskpane。

`--pairing 1` 仅在部署后的显式检查中使用。它从 `PPT_AGENT_RELEASE_PC_TOKEN` 读取专用测试 PC 身份令牌，使用固定 Office Origin 建立两条临时 WebSocket 连接，完成 v2 创建、协商、领取、批准，确认两端协商 `presentation.v1`、`presentation-attachments.v1`、`presentation-assets.v1`、`presentation-remote-images.v1`、`presentation-asset-rights.v1`。它对每种能力分别用随机短数据核对请求转发、分片响应及完成回执，随后关闭连接。不要在命令行传令牌，也不要使用日常用户令牌。检查失败时不会打印令牌、配对码或会话凭证。此测试仅覆盖 Relay 协议、能力路由和该测试身份的认证路径；不调用真实 PC 客户端的附件或图片操作，也不代表 PowerPoint 宿主已经通过验收。
