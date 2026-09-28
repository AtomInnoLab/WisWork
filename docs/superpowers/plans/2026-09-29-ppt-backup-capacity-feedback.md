# 保存点容量与恢复反馈

原方案阶段 5 要求面向用户的错误和恢复动作。基线 f333bf08，隔离实现分支。

根因：UI 写死活动保存点上限8而PC当前为16，且整页变更使用原包/源包两份。PC quota_exceeded 被页面与共享图表备份读取层泛化，提案过滤后只剩通用写入失败。

实现：UI显示实际活动包数量，不推断连接PC的容量上限；解释整页占两份包、释放已结束记录可恢复容量。两个已有备份响应边界仅在 begin 收到精确 quota_exceeded 时抛 presentation_existing_backup_capacity；其余错误保持既有过滤。提案控制器仅额外允许这一稳定错误码，Agent确认错误呈现中文释放/重试动作，不泄露响应原文且不自动重试。工作台对容量、备份缺失和各类现稿冲突显示明确恢复提示。

1. 核心错误传播：presentation-existing-page-editing.ts、presentation-chart-backup.ts、proposal-controller.ts、use-office-agent.ts及对应测试。失败先行核对真实响应容量拒绝发生在宿主写前，提案/会话返回稳定用户提示，未知错误仍隐藏；定向测试、类型、lint/格式后提交。
2. UI与控制器：presentation-changes-card.tsx、presentation-changes.ts及对应测试。去除假定分母，准确说明双包，已知错误映射与未知内容不泄露。
3. 独立完整复核、Office全套测试及更新进度。无存储/备份限额改变，无新目的地，不把验证计为真实宿主验收。
