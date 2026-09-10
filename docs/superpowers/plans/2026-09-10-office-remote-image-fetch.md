# Office 远程图片获取实施计划

基线提交：`865b8f5`，分支：`codex/fix-office-remote-consent`。规范：`docs/superpowers/specs/2026-09-10-office-remote-image-fetch-design.md`。

## 交付 1：Relay 受控图片端点

文件：`services/wiswork-relay/src/lib.rs`、`services/wiswork-relay/tests/relay.rs`、`services/wiswork-relay/deploy/nginx-location.conf`、`services/wiswork-relay/deploy/nginx-http-limits.conf`、`services/wiswork-relay/README.md`、必要的 Cargo lockfile。

先写失败测试，证明未认证请求、私网 DNS、未固定解析、重定向逃逸、错误 MIME、超限、超时和并发不能通过，并证明合法公网 JPEG/PNG 可以返回。最小实现新增精确认证 POST 路由和逐跳固定地址下载器，复用现有 OIDC 鉴权，不接触 WebSocket session/store。运行 Relay 完整测试、clippy 和部署配置断言，形成独立提交。

## 交付 2：PC 远端优先下载

文件：`apps/shell/src/main/office-retrieval-proxy.ts`、`apps/shell/src/main/index.ts`、对应 Shell 图片连接/管线/代理测试及安全文档。

先写失败测试，证明生产 wiring 目前没有远端下载器，以及远端成功时不调用本地直连；覆盖鉴权、取消、响应 MIME/体积和远端失败后的严格本地回退。实现编译固定端点和有界远端下载函数，将其注入现有每会话搜索代理；不改搜索账本和插件协议。运行 Shell 相关完整套件、类型与生产构建，形成独立提交。

## 交付 3：集成验证与发布准备

检查 Relay 端请求/响应与 PC downloader 接口完全一致，增加必要的端到端测试，执行全仓 `npm test`、`npm run typecheck`、lint、格式、diff，Relay `cargo test --locked`、clippy 和 release build。进行独立安全审查并修复 Critical/Important；记录部署、健康检查、真实取图冒烟和回滚步骤。本交付不在用户再次明确要求前升版或部署生产。

### 验证记录（2026-09-10）

- Relay：31 个单元测试、54 个集成测试、doc tests 全部通过；`cargo clippy --all-targets -- -D warnings` 通过。
- PC 图片链路：3 个目标测试文件共 78 个测试通过；Shell 全套 613 个通过、11 个跳过；Shell 类型检查和生产构建通过。
- 全仓：类型检查、格式检查、`git diff --check` 通过；ESLint 0 错误（13 个既有 warnings）。
- 全仓测试唯一失败为不相关的 PPTX 性能用例在并行负载下用时 5.14 秒、超过 5 秒阈值；该用例随后独立重跑用时 3.45 秒并通过。
- Relay release build 与最终独立跨组件安全审查在本计划完成前执行并记录。
