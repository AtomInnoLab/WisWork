# Office 多步骤 Agent 请求超时修复

## 诊断证据

- PC 0.6.60，Taskpane build `1e2566080877`。
- Trace `c3e3a9e2-92ae-451e-9e8e-4a424361dc3d`。
- PC 任务从 `1789004096987` 持续到 `1789004376925`，状态为 cancelled。
- Taskpane 在 `1789004376855` 上报 request_timeout，总时长 280074 ms。
- 最后一次页面写入在 `1789004373556` 已得到 proposal_applied / mcp_tool_completed，约 3 秒后整轮被取消。
- 配套链路仍有 Taskpane Relay session 290s、Relay 300s、PC watchdog 305s 的绝对请求上限。

根因：Enhanced 的单条传输请求承载整轮模型与工具循环，却沿用单次模型响应的固定时限。有工具进展仍会被截止时间取消。

## 实现

1. Taskpane：Enhanced 请求按有效工具活动、工具执行与非空内容增量续期 280s 无进展计时器，同时限制整轮最多 30 分钟。
2. 心跳与未知事件不续期；Standard 响应仍采用原 280s 绝对时限。
3. v2 agent.v1 的 Taskpane Relay session / Relay / PC 上限分别为 30 分钟 + 10 / 20 / 25 秒，预留取消传播时间。
4. Legacy、检索请求、授权有效期及字节/事件数量限制继续独立生效。这里的请求上限不是延长授权；现有 Enhanced session statement 仍有自己的 15 分钟有效期。
5. 不自动重放页面写入，以免超时边界发生重复修改。

## 回归验证

新增测试先在旧实现失败，再在修复后通过：

- Taskpane：持续工具进展可越过 280s；停止进展、仅心跳、整轮绝对上限仍会取消；Standard 不续期。
- PC：同一 Enhanced 请求在 360s 后仍能转发工具结果并完成。
- Taskpane Relay session：请求超过五分钟不取消，最终 watchdog 取消一次并保留配对。
- Relay：真实 WebSocket 请求超过缩短后的旧上限仍能转发 chunk，最终新上限通知双方，之后可发送下一请求。

相关完整测试通过：Office 813 项，PC 524 项（11 项条件跳过），Relay 22 项单元测试 + 50 项集成测试。
Office / PC 类型检查与 Relay clippy（-D warnings）通过。
修改文件的 ESLint、Prettier、git diff --check，以及 Taskpane 生产配置构建和 PC 构建通过。全仓 lint 首轮只发现新增测试中的一处 prefer-const 错误（已修复并重验），其余为 13 条既有 warning。

## 边界与发布

后续针对单次工具失败的复核与修复见 [PC 0.6.61 配套发布](2026-09-10-office-single-tool-release.md)：已复现页数缓存未更新、非连续原型占位死锁、插图嵌套请求争用和 PC 工具帧上限不一致。原始诊断没有完整参数，不能声称每条历史错误均已精确归因；尤其脚本回读校验与计划解析错误仍需要更具体的现场证据。

此次改动需同时发布 PC、Taskpane 和 Relay。仅刷新现有 Taskpane 不能更新 PC/Relay 内的旧 watchdog。
真实 macOS PowerPoint 长时间制作用例仍需在配套更新后验收；自动测试不替代宿主端验收。
