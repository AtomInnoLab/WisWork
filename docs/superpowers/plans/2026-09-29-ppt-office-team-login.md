# Office 团队登录与连接实施计划

遵循原方案 P2 团队身份和评论链路；复用已存在 WisWork 网关兑换/刷新协议及 Relay OIDC 校验，不新增服务器认证协议。

1. 将 OAuth 状态机提取为共享核心，新增浏览器入口、WebCrypto 和仅内存凭证存储，保留桌面兼容性。
2. 使用 Office DialogOrigin 1.1 对话框：同源启动页、明确注册的 HTTPS 回调、来源和状态校验、取消及迟到回调隔离。
3. 独立 presentation-team.v1 连接，令牌只在该连接传递；账号或连接失效立即清除团队内容及待确认操作。
4. 配置、CSP、构建页面和工作台连接；未配置网页登录 clientId 与回调时禁用，不能默认复用桌面注册。
5. 独立安全审查、定向测试、相关工作区回归、类型及构建，更新进度证据。

边界：不注册 IdP、不部署、不声称模拟测试证明实机登录。当前等待已注册的网页 clientId 和 HTTPS 回调；实际网关 CORS、IdP 配置、PowerPoint 对话框和生产 Relay 身份链需实际验证。

参考：[Office 对话框认证](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api)。现有网关未提供 PKCE verifier 合约，不自行新增参数。
