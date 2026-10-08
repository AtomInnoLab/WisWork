# PC 0.6.62：图片链路修复与 Taskpane 0.3.33 发布

## 已确认范围

用户确认：优先完成图片链路修复，再升级 PC、提 PR、更新本机 Taskpane。DESIGN.md 显眼预览入口和直接 PC 编辑/保存同步的新方案暂停，不新增跨端文件协议，不部署 Relay，不自动合并或创建安装包发布 tag。

沿用隔离分支 `codex/fix-office-review-recovery`，保留并一起提交前轮已验证的合同审查恢复、合法工具响应大小、样式细分诊断与图片错误展示修复。部署使用已有本机 Nginx 静态目录，按原有访问范围发布。

## 图片实现单元及验收

1. `packages/ai-search/src/index.ts`：逐项过滤坏候选，不丢弃同源合法结果；保留 HTTPS/来源安全校验，全坏响应仍报错。三个 provider 的混合候选测试先 RED 后 GREEN。
2. `office-image-handoff.ts`、Shell `index.ts`、`office-retrieval-proxy.ts`：复用现有有界编码器，将缩放放在全尺寸转码超限拒绝之前；只宣传可解码的 PNG/JPEG。保留原图 2MiB、8192px/16Mpx、HTTPS/DNS/redirect 和工具输出预算。组合测试覆盖下载→实际生产 normalizer 绑定→handoff、超限和不支持格式。
3. `browser-powerpoint-import-media-adapter.ts`：插入只执行一次，校验最多回读三次；持续不符/取消继续原有清理。测试暂时旧值恢复、永久不符和取消，不放宽 ID/类型/几何检查。
4. 图片来源授权过期：每个 Relay 客户端独立维护有界 100 条记录和原查询；下载前重跑原查询，只有新结果再次包含同一 URL 才续期。断开、撤销或重置时清空记录，并以 generation 隔离迟到的搜索、下载和编码结果。未搜索/被移除/取消/旧会话场景不得下载或恢复授权，不绕过来源授权，不把任意 URL 加入白名单。

## 验证与发布

- 各单元独立审查，最终运行完整 `npm test`、`npm run typecheck`、lint、格式检查和 PC/Taskpane 生产构建。
- 版本只调整 PC package/lock 和 Office manifest/构建断言。当前上游 main 内容与旧分支 HEAD 一致（上一 PR 被 squash），提交后将新提交移至最新 main，避免 PR 重带旧提交。
- 使用已有 PR 模板提交到 main。PC 安装包在合并后的正式发布流程生成，本机无法代表用户 Mac 已升级。
- Taskpane 先备份 `/var/www/wiswork-office-addin`，先上静态资源后切换入口与 manifest，保留旧 assets/components；核对公开入口、版本、构建 ID 和产物摘要。失败恢复备份入口，无数据库变更。
- 不把 mock/编码器适配测试描述为真实 macOS PowerPoint 实机验收；旧事故 URL/实际属性已脱敏，保持归因边界。

## 发布前验证记录

- 最终完整 `npm test` 通过，其中 Office Add-in 868 项、PC Shell 569 项通过，PC 的 11 项按现有条件跳过。
- 最终全仓 `npm run typecheck` 通过。
- 全仓 lint 为 0 errors、13 条既有 warnings；变更文件格式检查及 `git diff --check` 通过。
- PC Shell 0.6.62 生产构建通过；Codex release、rollout 和七宿主 golden 检查通过。
- 图片相关单元先 RED 后 GREEN。独立审查发现并修正了图片来源记录跨客户端复用的问题，最终审查无剩余 Critical/Important。
- Taskpane 部署后的构建 ID、公开资源校验与回滚备份位置在本次 PR 的部署记录中补充。
