# PPT Agent：单页替换正式提交与整页撤销

日期：2026-09-23。基线 `4011fef3`，分支 `codex/ppt-agent-implementation`。

## 1. 对应原方案

按原方案 §6.4、O3、O5、§21，继续完成单页重做的写入事务和整页撤销。原方案不变。本批补正式提交及从备份恢复；真实 Office 写后验收、内容证据 QA 和 P0 仍为进行中。

## 2. 完成内容

- 新增确认式 `commit_presentation_page_replacement` 和 `undo_presentation_page_replacement`，预览包含业务页、原页、修订页与备份。
- 提交前重新核对 PC 备份、父子产物及完整回执；先保存 `commit_pending`，核对双页内容与顺序后删除原页，回读再保存 `applied`。
- 撤销先保存 `undo_pending`，插入原始备份，回读唯一恢复页 ID 并保存 `restore_inserted`；验证原页已恢复后才删除修订页，回读后保存 `undone`。
- 文档设置使用同一 envelope 保存事务和回执覆盖映射。提交时父任务标记已取代、子任务接管业务页；撤销时恢复父任务映射到新恢复页 ID。避免映射丢失后误追加整套。
- 删除前预检映射容量；保留旧裸事务记录读取，冻结提交中的相关回执，并沿用 CAS、串行保存、失败回滚和文档身份校验。
- 派生任务仅在已有完整、摘要和页面身份一致的回执时可 prepare；拒绝已取代任务并保留旧缓存。提交后显式 prepare 子任务，撤销后 prepare 父任务，继续编辑与 QA。
- runtime 接通工具、局部 QA 失效和状态通知；隐藏未决破坏性操作或已被取代缓存的导入/QA 状态，避免误报当前页面已验收。
- `commit_pending` 已完成删除时只补状态；`restore_inserted` 已完成恢复和删除时也只补状态。未知插入归属、内容漂移、文档切换或确认过期均停止自动处理。

状态：`staged → commit_pending → applied → undo_pending → restore_inserted → undone`。暂存撤回仍走原有路径。

## 3. 验证证据

- 全仓 `npm test` 退出码 0：6194 项 Vitest、2 项 Node 脚本测试及 53 项 Rust 测试通过，另有 6 项 Vitest 跳过。日志 `/tmp/wiswork-commit-full-test.log`。这些不是 20 项专业任务验收。
- 全仓类型检查、Lint（0 错误、9 条既有警告）、格式及相对基线格式检查、依赖许可、diff 检查通过。
- Office Addin 和 Shell 构建通过；Addin 保留大于 500 kB 的产物体积告警。
- 独立审查覆盖事务、映射、宿主恢复和缓存路径，验证 97 项定向测试及 1 项跨层测试通过，未发现阻塞问题。根检查补修 undo 的 QA 影响目标，断言先失败后通过，34 项定向测试通过；复审确认撤销后的父页历史 QA 不会被视为当前通过。
- TDD 日志包括 `/tmp/journal-red*.log`、`/tmp/wiswork-commit-integration-red.log` 与 `/tmp/replacement-undo-impact-{red,green}.log`；全仓及构建日志为 `/tmp/wiswork-commit-*.log`。

跨层测试使用真实 PptxGenJS 父子产物、PC 备份服务、文档 settings binding 和正式宿主适配器；底层 Office 上下文为模拟。原宿主页含人工修改，测试确认撤销恢复该原始备份字节。分别注入删除原页后保存 applied 失败、删除修订页后保存 undone 失败，重启 PC 与插件后完成回执，不重复插入/删除。非目标页顺序和映射保持一致。这不是实机 Office 验收。

## 4. 边界与风险

- 包指纹保守比较 OOXML 条目。真实 Office 导入可能规范化关系、ID、元数据而触发冲突，需实机验证，不能将模拟上下文成功当成宿主兼容成功。
- Office.js 没有跨导出、settings 与删除的原子 CAS，最后读取至写入间仍有共同编辑窗口；需要真实宿主并发验收。
- 未知插入 ID 不自动认领或再次插入；备份不可用、页面内容/位置被修改时保留现场。人工解除未决记录的产品流程仍未开放。
- 当前是一份最近事务的整页撤销，不是任意现稿、多级历史或通用 ChangeSet 撤销。新事务仅在上一份 discarded/undone 后开始。
- 回执覆盖最多 32 个键、100000 字节，事务最多 192 KiB；已有 PC 备份和包解析限额继续适用。已被取代的任务 ID 不复用。
- 新 envelope 不能由旧插件继续操作；回滚必须保留备份、事务及覆盖映射。旧文档无需预迁移。

## 5. 部署与下一步

本批仅在实现分支提交，未合并、推送或部署；PC/Relay/Manifest 无新增协议变更。

下一步按原方案补替换后页面 revision 与截图/结构/内容证据 QA 的验收衔接，再开展真实 PowerPoint 保存、关闭、重开、可编辑性和并发修改验证。20 项专业任务仍未执行，O3/O5/P0 不据此宣布通过。
