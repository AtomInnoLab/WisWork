# PPT Agent 发布兼容矩阵（工程证据）

2026-10-01 本机补充：发布预检把 `taskpane.html` 的全部 `modulepreload` 引用纳入校验，与团队登录两页一致；缺文件或外域/非哈希路径均拒绝。预检还从发布包内 JS/MJS 的相对动态/静态导入及根路径或相对 Worker URL、CSS 的根路径或相对 URL 收集运行时依赖，要求引用文件存在且符合哈希命名。缺失分包、Worker 和图片用例均被拒绝，预检测试 18/18 通过；设置本机 HTTPS 测试 Origin 后构建的 17 个发布文件通过完整检查。未设置 `VITE_WISWORK_ADDIN_ORIGIN` 的普通本地构建不会生成发布 Manifest，不作为完整发布包。

2026-10-01 本机补充：生产 Taskpane 的 `version.json` 请求失败、超时或元数据无效时进入可重试的版本未核实状态，文档工具和 PC 连接入口均不开放；开发服务器仍允许缺少该文件。构建版浏览器冒烟首次注入 HTTP 503，核对只显示版本重试入口，再恢复版本响应并完成真实 Relay/PC 业务链。此为本机生产构建和模拟 Office.js 宿主的版本门禁证据，不能替代部署站点或真实 PowerPoint 混合版本验收。

依据[完整方案](./wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md) §10 O6。此表描述当前协议实现和自动化证据，不代替 Windows/Mac PowerPoint Desktop/Web 的发布验收。

本机 `node tools/ppt-agent-electron-real-relay-smoke.mjs` 还验证 Chrome Taskpane 发起真实 Relay 配对、从 Electron PC 读回已持久编译的项目，并在传输中断后续接会话、重新读取项目；Taskpane 重开后重新配对仍能恢复同一项目。浏览器使用 Office.js 测试宿主和协议帧传输桥；这项结果不覆盖发布域名的网络连接或真实 PowerPoint 宿主。

| Taskpane / Relay / PC              | 当前行为                                                                                                                                                                                                                                                                                   | 工程证据                                                                                                                                 | 发布状态                                                |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------- |
| v2 / v2 / v2，协商 PPT 能力        | 按交集启用 `presentation.v1`、附件、图片和权利声明能力；未协商能力在发送前拒绝                                                                                                                                                                                                             | `apps/office-addin/tests/relay-session.test.ts`、`services/wiswork-relay/tests/relay.rs`、`apps/shell/tests/office-relay-client.test.ts` | 工程路径已覆盖；真实宿主待验收                          |
| v2 / v2 / v2，PC 只提供 `agent.v1` | 保留普通 Agent；PPT 专用请求在 Taskpane 侧拒绝                                                                                                                                                                                                                                             | `apps/office-addin/tests/relay-session.test.ts`                                                                                          | 工程路径已覆盖；真实旧 PC 构建待验收                    |
| v1 / v2 / v2                       | Relay 仍接受 v1 基础配对；仅提供旧版 `agent.v1` 协议                                                                                                                                                                                                                                       | `services/wiswork-relay/tests/relay.rs`、`apps/office-addin/tests/relay-session.test.ts`                                                 | 工程路径已覆盖；真实旧 Taskpane 构建待验收              |
| v2 / v2 / v1                       | 旧 PC 用 v1 认领当前 v2 验证码时，Relay 向 PC 返回版本错误，并按配对 ID 向 Taskpane 发送版本不兼容事件；新版 Taskpane 立即提示升级 PC，旧 PC 自身仍显示原有通用协议错误。未知验证码和不相干邀请不触发该事件；若旧 PC 未实际认领，Taskpane 仍只能按原配对期限等待。生产版新 PC 使用 v2 协商 | `services/wiswork-relay/tests/relay.rs`、`apps/office-addin/tests/relay-session.test.ts`、`apps/office-addin/src/App.tsx`                | **发布阻塞：必须先升级 PC；旧二进制不能原地获得新提示** |
| v2 / v1 / 任意 PC                  | v1 Relay 不支持 v2 配对；收到精确旧版协议错误时提示升级，不会静默降级                                                                                                                                                                                                                      | `apps/office-addin/tests/relay-session.test.ts`                                                                                          | **发布阻塞：必须先升级 Relay**                          |

构建版本探测、Manifest、能力协商分别解决静态资源、协议能力和会话入口的问题；三者不能互相替代。发布时应先部署支持 v2 的 Relay 和 PC，再发布 v2 Taskpane，并保留旧 Taskpane 的回滚包。`version.json`、`taskpane.html` 与其哈希资源需原子发布；真实 Win/Mac/Web 混合版本与端到端冒烟仍待执行。

本机跨运行时门禁 `node tools/ppt-agent-electron-real-relay-smoke.mjs` 已覆盖真实 Electron PC 与 Rust Relay 的三次进程会话：首次完成八页编译、逐页生产、PPTX/PDF 读取和 TXT/PNG 附件上传回读清理，并在另一个项目保存计划、开始一项待执行的八页生产任务；PC 正常退出后以相同用户数据目录重启并重新配对，继续执行待处理任务，逐页检查 PPTX、导入来源和 PDF，再只读检查已完成任务的全部交付摘要。第二进程另起独立八页后台任务，第一页完成、第二页进入编译时强制退出；第三进程检查持久 `interrupted` 回执，显式续跑后确认第一页未重编、第二页重试且全部八页完成。它验证本机 PC 进程恢复读取、待执行任务续接和部分完成后台任务恢复，不代表系统断电时磁盘耐久性、部署后的 Relay/PC/Taskpane 混合版本或 PowerPoint 宿主恢复验收。

灰度构建可设置 `VITE_WISWORK_PRESENTATION_ROLLOUT_PERCENT=0..100`（默认 100）。Taskpane 按已保存文稿的 Office 地址只读分桶；部分灰度中未保存或未命中的 PowerPoint 文稿在启动 Agent 前显示不可用状态。比例随新构建发布，回滚需要保留并切回旧构建及 Manifest。该开关只控制 PPT Agent 入口。可另设 `VITE_WISWORK_OFFICE_DIAGNOSTIC_SAMPLE_PERCENT=0..100`（默认 100），按运行 trace 对脱敏失败诊断远程抽样，本地诊断仍完整；抽样不是健康指标验收，也不能代表真实宿主兼容或部署后全链路冒烟已完成。

发布预检命令（部署前先用构建产物执行，部署后再加 `--deployed 1`）：

```bash
node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist
node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist --deployed 1 --rollout-percent 25 --diagnostic-sample-percent 10
PPT_AGENT_RELEASE_PC_TOKEN='<release-test-PC-token>' node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example --dist apps/office-addin/dist --deployed 1 --rollout-percent 25 --diagnostic-sample-percent 10 --pairing 1
```

部署后预检要求显式填写本次发布的灰度比例和诊断采样比例，并与 `version.json` 中的构建元数据逐项核对；示例的 25% 和 10% 需替换为实际发布值。预检核对 Manifest 各 Office 资源地址及其本地文件、HTML 引用的样式文件、`version.json` 与编译入口的一致性、JS/CSS/worker 等运行时资源的哈希式文件名、除官方 Office.js 外无额外脚本、无源码映射、Relay `/office-relay/health` 的精确响应；本地构建的全部发布文件（含 CSS、worker、图标和 Manifest）形成大小与 SHA-256 清单。部署后逐一限量读取线上文件并比对完整字节，同时要求 `version.json`、`taskpane.html`、Manifest 与固定图标响应 `Cache-Control: no-store`，带哈希资源响应至少一年的 `immutable` 缓存策略；缺失、超量、篡改或缓存策略错误均失败。任何失败均以非零退出码阻断继续发布。CI 会构建真实 Taskpane 并执行产物预检，以及使用 LibreOffice 检查 PPTX 包回读。此命令不替代 v1/v2 配对协议测试（CI 的 Relay/Office 测试）及真实 PowerPoint 的附件、图片、页面写入、截图、恢复冒烟。诊断协议新增 `run`/`passed` 与 `agent_run_completed`；须先部署支持该事件的 Relay，再发布新版 Taskpane。事件只证明 Agent 流程结束，不代表 PPT 交付完成；真正完成率仍需 20 项任务与宿主证据。先保留旧版回滚包；发布顺序仍为 Relay、PC、Taskpane。

`--pairing 1` 仅在部署后的显式检查中使用。它从 `PPT_AGENT_RELEASE_PC_TOKEN` 读取专用测试 PC 身份令牌，使用固定 Office Origin 建立两条临时 WebSocket 连接，完成 v2 创建、协商、领取、批准，确认两端协商 `presentation.v1`、`presentation-attachments.v1`、`presentation-assets.v1`、`presentation-remote-images.v1`、`presentation-asset-rights.v1`。它对每种能力分别用随机短数据核对请求转发、分片响应及完成回执，随后关闭连接。不要在命令行传令牌，也不要使用日常用户令牌。检查失败时不会打印令牌、配对码或会话凭证。此测试仅覆盖 Relay 协议、能力路由和该测试身份的认证路径；不调用真实 PC 客户端的附件或图片操作，也不代表 PowerPoint 宿主已经通过验收。

实际 PC 业务冒烟使用独立命令。先在发布测试 PC 上登录 WisWork，并准备已存在的测试演示文稿文档 ID 和项目 ID，以及该文档中一份已解析文本和一张已规范化图片的附件 ID；将它们作为环境变量传入，避免放进命令行参数。命令显示一次性配对码，操作员在该 PC 输入并批准；随后它向真实 PC 的 `presentation.v1` 与 `presentation-assets.v1` 服务发送只读状态、附件清单、文本读取与图片读取请求，核对项目 ID、页数、附件身份、图片字节摘要、会话身份与分片顺序。文本和图片 ID 未设置时只检查状态与清单，并明确报告未检查对应内容。失败以非零退出码结束，不打印 PC 返回的项目或附件内容。

```bash
PPT_AGENT_SMOKE_RELAY_ORIGIN='https://relay.example' \
PPT_AGENT_SMOKE_DOCUMENT_ID='<test-document-id>' \
PPT_AGENT_SMOKE_PROJECT_ID='<test-project-id>' \
PPT_AGENT_SMOKE_TEXT_ATTACHMENT_ID='<text-attachment-sha256>' \
PPT_AGENT_SMOKE_IMAGE_ATTACHMENT_ID='<image-attachment-sha256>' \
node tools/ppt-agent-pc-business-smoke.mjs
```

这项检查需要真实 PC 在线与人工完成一次配对；仓库中的模拟 Relay 测试只验证命令自身的协议和错误门禁。完整配置两个附件 ID 且通过后，可证明部署后的配对、PC 项目状态及既有文本/图片资产读取。页面写入、截图、恢复和 PowerPoint Desktop/Web 仍须按 O6 矩阵逐项执行，不能由此命令宣称整体冒烟或专业任务通过。

专用发布测试文档还可显式设置 `PPT_AGENT_SMOKE_PRODUCTION=1`。命令在给定项目名前缀后追加随机后缀，创建新的八页测试项目，保存计划、逐页制作并核对每页原生 PPTX、导入来源和八页 PDF；结果打印实际项目 ID 以便追查。该项目会保留在测试 PC 的本机项目库中，因此只在专用测试文档使用；不设置该变量时不会创建项目。本机真实 Electron PC + Rust Relay 冒烟已覆盖从空项目开始的这条链路，发布环境仍需单独运行。该路径不写入当前 PowerPoint 宿主页面，也不能替代真实宿主截图、编辑、保存重开和恢复验收。

```bash
PPT_AGENT_SMOKE_RELAY_ORIGIN='https://relay.example' \
PPT_AGENT_SMOKE_DOCUMENT_ID='<dedicated-test-document-id>' \
PPT_AGENT_SMOKE_PROJECT_ID='release-test' \
PPT_AGENT_SMOKE_PRODUCTION=1 \
node tools/ppt-agent-pc-business-smoke.mjs
```

专用测试文档可设置 `PPT_AGENT_SMOKE_UPLOAD=1`（不同时设置两个既有附件 ID），在同一配对会话中上传带随机内容的小型 TXT 和 1×1 PNG、读取解析文本及规范化图片、核对图片摘要，然后按精确 SHA-256 附件 ID 删除两个测试附件。命令先检查随机 ID 不存在，只清理本次创建的 ID；任一步失败也尝试清理。若清理失败，命令输出需人工检查的附件 ID 并返回非零。该选项会短暂修改 PC 的测试文档附件目录，**只对专用发布测试文档使用**。仓库 CI 的 Electron 图片冒烟已使用相同 PNG 生成方式经过真实本机解码器和 PC 附件服务的上传/回读/删除；发布环境仍须运行上面的配对命令。PowerPoint 宿主写入、截图和恢复仍需单独验收。

```bash
PPT_AGENT_SMOKE_RELAY_ORIGIN='https://relay.example' \
PPT_AGENT_SMOKE_DOCUMENT_ID='<dedicated-test-document-id>' \
PPT_AGENT_SMOKE_PROJECT_ID='<test-project-id>' \
PPT_AGENT_SMOKE_UPLOAD=1 \
node tools/ppt-agent-pc-business-smoke.mjs
```
