# PC 0.6.61：Office 单次工具失败与配套发布

## 目标与边界

修复可复现的连续复制误拒绝、Enhanced 插图下载与整轮请求争用问题；将已有多步骤超时修复一起发布 PC PR、Taskpane 和本机 Relay。
不重放写操作、不绕过设计契约/确认/回读验证、不扩大 Relay 单请求并发或工具输入 256 KiB 上限，不延长授权有效期。

## 架构与实现单元

1. `powerpoint-skill.ts`：只有复制的新页 ID 经准确验证成功后才更新已知页数；兼容专用与旧声明式入口。测试先重现“复制页 0 后页 1 越界”，再验证连续复制、真实越界和契约门禁。对非连续原型，只允许在尾部追加到达后续原型必需的占位页，确认时复核当前页数与契约；占位页保持 dirty，但不算 built / review pending，后续仍须填充验收。现代契约下的旧声明式复制明确引导到专用工具，避免把来源页错记为目标页。
2. `office-codex-proxy.ts`、PC 原生图像编码与 Taskpane 插图链路：复用 PC 已验证检索能力预取图片，经有界编码后随本次工具调用私有传递；Taskpane 在可信远端入口拆出，单次注入插图执行。模型公开输入继续只提供 URL；私有数据不得来自模型、不得写入全局未绑定缓存。测试要求整轮请求活跃时能插图，下载/校验失败无 Office 写入，取消与批准检查继续生效。PC `office-relay-client.ts` 的工具帧改用既有 256 KiB 输入上限加控制元数据预算，其余控制帧仍限 16 KiB；已下发写入的传输故障保留原错误传播，不能误报为预取失败/未修改。
3. `office-diagnostics.ts` 与 `relay/session.ts`：本地保留图片故障细分码，发送给旧 Relay 时映射至已有协议词汇，避免误显示为 Office 写入错误或产生诊断上传失败。新增映射测试先失败再通过。
4. PC `package.json` / lock 升为 0.6.61；Office manifest / 对应构建测试升为 0.3.32。提交并向 main 提 PR，不预建发布 tag、不自动合并。

## 验证与证据约束

- 诊断中的 1–2 ms 图片失败与同步 `relay_busy` 路径一致；本地错误码过滤又会将 `image_fetch_unavailable` 折叠成 `office_write_failed`。
- 诊断记录删去了原始参数，不能据此断言每一条历史复制、脚本验证或计划错误都已精确定位。对无法复现的分支明确保留边界。
- 每单元定向 RED/GREEN，最终完整 `npm test`、`npm run typecheck`、lint、构建；Relay test、clippy 与 release 构建。独立审查跨端图片数据边界及兼容性。
- 真实 macOS PowerPoint 复制/插图/长流程验收不能由 mock 测试替代。

### 最终验证记录

- 全仓 `npm test` 通过；最后两个占位页保护测试加入后重跑完整 Office 套件，834 项通过。PC 套件 547 项通过、11 项条件跳过。
- 全仓 `npm run typecheck` 通过；lint 为 0 error、13 条既有 warning。PC 生产构建、变更文件 Prettier 与 diff 检查通过。
- Relay 22 项单元测试、50 项集成测试、clippy `-D warnings` 与 release 构建通过。
- 七宿主 golden、品牌、快速迭代隔离与 Codex 发布/rollout 策略检查通过。
- 独立审查覆盖跨端图片边界和占位页门禁；已修正“远端写入连接失败被归类为下载失败”的审查发现。
- Relay 已按上述副本验证流程更新，公开健康检查为 `ok`，数据库仍为服务用户 `0600`。受保护回滚备份位于 `/var/backups/wiswork-relay-20260910-Q6r0Ni`。

## 部署与回滚

- 本机 Relay `/opt/wiswork-relay/wiswork-relay`，服务 `wiswork-relay`；先验证 release 二进制及持久化回归，备份二进制与受保护配对数据库，在 staging 副本验证启动，再替换并检查本地和公开健康接口。无数据库 schema 改动。
- Taskpane `/var/www/wiswork-office-addin`，产物构建使用生产 origin 与提交 build ID。先备份、复制静态资源再替换入口和 manifest；不删除无关 components 或旧缓存资源。
- 失败则恢复已记录的二进制/静态入口；保持配对数据库不变，恢复旧 Relay 前核实 schema 兼容。部署记录包含备份位置和产物校验值。
- PC 安装包须在 PR 合并后由正式发布流程生成；本次本机服务更新不等于用户 PC 已升级。三端 timeout 修复全部生效需新 PC。
