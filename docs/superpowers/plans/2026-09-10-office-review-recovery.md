# Office 制作校验与审查恢复

## 证据与边界

用户本轮 trace `9be0a91c-a163-4557-ab65-018db66c5755`、PC 0.6.61、Taskpane `f8c2bd17cc7f`。样式脚本在 execute 成功后两次 verify 失败；三页制作后重复收到 review_required；随后以 producing 提交 plan_deck 被拒绝。最终工具启动后约 265ms PC 任务取消，Relay 同秒记录 office.cancel 的 invalid_session。

原始内容已脱敏，不能还原具体字体、原始图片 URL 或最终响应字节数。已复现独立缺陷：PC 将合法大于 16KiB 的工具结果误判为控制帧；规划 schema 允许 producing/verified 而执行端拒绝；重复 ready 提交清空生产进度；截图后纠正消息错误引导 verify_slides 而非登记 review。

## 修复单元

1. PC Relay 客户端：只让 relay.tool_result 使用已有工具帧上限；保留 272KiB 硬上限、控制帧 16KiB、会话/调用身份匹配。测试 32KiB DESIGN.md 与 256KiB 输出、后续调用，以及恶意/超限控制帧拒绝；不扩展名义 16MiB 协议预算。
2. PowerPoint 合同与审查：模型输入只公布 draft/ready；错误返回 pendingReviews、acceptance_ids 和下一步；相同 ready 计划重发保持当前 producing/verified、已生产页与待审记录。实际改稿仍重新校验并重置生产状态。review 必须有当前截图，禁止修改合同状态冒充验收。
3. 样式回读：校验整个有界实际文本，保留 mixed/unknown 值而不是强制 false/0；验证失败附加安全的 operation/property 位置，不导出正文、期望值或实际值。保留真正不匹配的拒绝，不扩大重试或重放写入。
4. 图片失败可见性：明确图片工具名称；PC 预取失败用已有安全错误码传入生命周期，Taskpane 显示并记录一次；复用现有校验/权限机制，不更改下载安全策略。设计契约安全错误码也在本地保留，旧 Relay 兼容映射。

## 验证与交付

各单元先 RED 再 GREEN；全量 Office、PC 相关套件、类型检查、lint/格式检查及构建，跨组件修改独立审查。保留真实 macOS PowerPoint 实机验收边界。

本轮请求是定位并修复；在独立分支完成代码和测试，不自动升级版本、改 PR 或部署服务。代码修复可按提交回退，无数据库/协议迁移。

后续用户已明确要求优先修复图片链路并升级 PC、提 PR、部署 Taskpane；本文件记录的修复随 `2026-09-10-office-image-release.md` 统一发布，发布边界以该后续计划为准。

## 实际验证结果

- `npm test` 全仓通过；其中 Office Add-in 865 项通过，PC Shell 554 项通过、11 项按现有条件跳过。
- Office Add-in 与 PC Shell 类型检查通过；全仓 lint 为 0 errors、13 个现有 warnings；所有变更文件格式检查及 `git diff --check` 通过。
- PC Shell 与 Taskpane 生产构建通过，仅产出本地构建，没有部署。
- 独立审查发现并修正完整 draft 被误识别为 ready 重放、旧 Relay 不接受新增诊断码两处问题；复审无剩余重要发现。另验证 verified 状态重放、真实改稿失效旧审查，保持安全门槛。
- 精确历史样式不匹配属性、图片失败来源、最终响应大小无法从已脱敏日志恢复；此次没有声称完成 Mac PowerPoint 实机验收。
