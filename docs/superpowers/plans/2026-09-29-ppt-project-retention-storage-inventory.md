# 项目保留期与统一删除：存储核对及实施边界

按原方案§13.1核对，基线93df802f。本文为下一实施单元的输入，尚未实现保留期或统一删除。

| 存储 | 归属 | 删除需要解决的问题 |
|---|---|---|
| projects/presentations/<projectHash> | 项目，内部绑定文档 | 计划/修订/生产/编译产物/来源审查/任务/反馈/资产事件；统一停止worker与禁止新写 |
| presentation-research/<documentHash>/<projectHash> | 文档+项目 | 正式记录、归档及删除回执，不能遗漏已有清理状态 |
| presentation-delivery-bundles/<documentHash>/<projectHash> | 文档+项目 | 完整包、上传中暂存及清单，和项目写锁协调 |
| presentation-page-backups/<projectHash> | 项目 | revision原页备份及暂存，删除前处理运行中的页修改 |
| presentation-preferences / presentation-comments / presentation-manual-observations | 文档+项目组合hash文件 | 包含本项目内容，需精确路径及内部归属核验 |
| presentation-teams | owner+document+project生成team ID | ACL和已发布内容，不能仅删除私有计划 |
| presentation-attachments/<documentHash> | 仅文档 | 目前无项目独占归属；建立显式引用/所有权，不能误删同文档其它项目资料 |
| presentation-acquisition-history/<documentHash>.json | 仅文档 | 含远程来源和附件结果；需按引用处理，不能把来源审查当保密审计 |
| presentation-existing-page-backups/<documentHash>及.released | 仅文档 | 修改事务仍引用原页；需项目关联与未决保护，不能为删一个项目清空整文档 |
| presentation-brand-kits | 用户全局共享 | 本项目引用解除；无授权不删除全局品牌资源 |

## 实施顺序

1. 持久项目生命周期(policy/revision/deletionId/state)，全入口统一在项目锁下校验，worker与长下载完成前二次校验；删除中的项目不能产生新记录。
2. 建立项目资源引用和历史存量归属核对，文档共享资料仅移除明确项目引用；未证明独占的资源不得删除。明确给出未清理项，不能将partial宣称complete。
3. 用户可见预览与确认后保存删除intent，分资源持久回执，失败可重开按实际剩余状态继续；所有文件操作拒绝symlink和目录越界。保留期到期走同一删除事务。
4. 持久、可导出的最小审计记录：动作/时间/本项目匿名标识/结果/回执摘要，排除正文、附件、原始URL、完整documentId。审计保留策略与项目内容保留策略分别明确。
5. 合成PC临时目录集成验证多项目共享、并发上传/worker、进程重开、断点删除、ACL、预算、路径攻击；不得清理真实用户数据。

该范围完成后仍需原方案宿主/专业门禁证据，不能用测试数量推升整体百分比。
