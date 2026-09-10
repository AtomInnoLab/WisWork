# Office 远程图片获取设计

## 目标

让已连接 WisWork PC 的 Office 插件在本机 Node 直连被代理、VPN 或网络策略阻断时，仍能可靠获取本轮图片搜索选中的 PNG/JPEG，同时保持来源授权、SSRF、重定向、体积和超时边界。

非目标：不替代图片搜索，不允许任意 URL 下载，不绕过 DESIGN.md、原型、截图或几何验收，不让 Relay 保存图片或文档内容，不修改 Office↔PC WebSocket 协议。

## 架构

PC 继续维护每个 Relay 会话独立、可过期的搜索结果账本；只有账本中的原图或同候选备用图才能进入下载。下载改为调用同一 Office TLS 域名上的认证 `POST /office-image-fetch`：Relay 验证 PC Bearer token 后，在服务端逐跳解析 URL、拒绝任一非公网地址并把已验证地址固定到该次 HTTPS 请求，禁用自动重定向并重新验证每一跳。Relay 以原始 PNG/JPEG 响应，PC 继续使用现有 nativeImage 解码、缩放和重新编码后才转发插件。

本机严格直连下载器保留为远程服务不可用时的受限回退，但单个候选共享总时限，避免原图和备用图串行消耗多个完整超时。远程端点不接受搜索请求，也不改变 Relay capability 协议。

## 全局约束

- 端点仅接受认证 PC；Office Origin 请求、缺失/超长 Bearer、鉴权失败均拒绝。
- 请求必须是精确 JSON `{ "url": "https://..." }`，请求体不超过 4 KiB；仅 HTTPS 默认 443、无凭据、无 fragment。
- 每一跳先解析全部地址；任一 loopback、private、link-local、carrier NAT、benchmark、documentation、multicast、reserved、IPv4-mapped 或其他非公网地址即拒绝。
- 连接使用 `reqwest::ClientBuilder::resolve_to_addrs` 固定到已验证地址；自动重定向关闭，最多三跳。
- 整体取图时限 15 秒；响应只允许 `image/png`、`image/jpeg`，声明或流式内容超过 10 MiB 即中止。
- Relay 施加独立全局并发上限和 Nginx 请求速率/请求体限制；不记录 URL、查询、token、响应体或图片摘要。
- PC 对 Relay 响应再次验证状态、MIME 和体积，并继续现有解码与 2 MiB Relay 传输上限。
- 失败仍返回稳定的 `image_fetch_unavailable`/`image_limit`/`image_mime_unsupported`，不伪造成功或写入文档。

## 失败处理与回滚

Relay 不可达、鉴权失败或远端取图失败时，PC 可尝试现有严格本地固定-IP下载；取消信号终止远端和本地请求。回滚顺序为 PC 回退到本地下载、移除 Nginx 精确 location、恢复旧 Relay 二进制；无数据库或文档迁移。

## 验收

- Relay 测试覆盖鉴权、精确请求、私网 DNS、DNS 固定、重定向逐跳校验、MIME、声明/流式体积、超时和并发。
- PC 测试证明搜索授权仍是前置条件、远端下载优先、原图失败后同候选备用图可成功、取消不回退、远端失败才使用本地严格下载。
- 完整 Rust/TypeScript 测试、clippy、类型、lint、格式与生产构建通过；独立安全审查无 Critical/Important 遗留。
- 真实部署后从 PC 网络完成“搜索→远端获取→nativeImage 归一化→Relay→PowerPoint 插入”冒烟；旧半成品不会自动修复。
