# 项目统一删除执行器

## 目标与架构

按原方案 §13.1 和保留期存储核对实现可预览、明确确认、可重开继续的本机删除。复用生命周期 intent/CAS/逐资源回执、真实 foreground/background drain 和既有项目锁；不另建任务调度器。共享、无独占归属证明和未决恢复资料保留，明确报告 partial，不伪造 complete。

删除先持久冻结，锁外取消并等待真实任务结束，再取得项目锁核验资源并清理。每个资源在独立治理目录保存原始归属与文件清单证明；原目录移入自有隔离区后逐项移除，断点使用同一 deletionId/证明继续，最后保存资源结果。项目正文最后清理，最小生命周期 tombstone 和匿名审计独立保留。

## 约束

- 本轮仅合成临时目录，不删除真实用户资料或部署。
- 原方案保持不变；全写入口接线/固定版本/drain gate 未完成前，不向实际普通 PC 调用开放执行入口。
- 用户批准的是明确项目 scope、删除范围和共享保留结果；计数为预览观察，冻结并 drain 后重新核验实际归属。不可把未知命名空间或新共享资源默默当作已批准独占资源。
- 只使用固定命名空间映射，不接受用户路径。拒绝目录越界、符号链接、身份替换和不匹配的 project/document。
- 普通版本不续租；删除只允许同一 intent 下通过自身 CAS 回执取得下一版本。等待锁的旧删除请求不能接纳其它执行器的新版本。
- 不删除属于其它调用的 staging，不自动重放未知写入。隔离区和证明具有确定性身份，source/quarantine 同时存在则拒绝而非猜测。
- 保留库存 32768 文件/32GiB、16 层扫描预算和原业务预算；新增治理证明单独有界，超限必须在文件删除前拒绝。
- 不宣称跨 OS 写进程原子隔离。已经开始的 OS I/O 允许真实完成；所有登记任务结束后才清理。
- 公开响应和审计只有有限状态/计数/匿名身份，不含路径、正文、原始 URL、内容摘要或原始异常。

## 单元 1：资源归属与隔离清理原语

文件：`apps/shell/src/main/presentation-project-deletion-resources.ts`，`apps/shell/tests/presentation-project-deletion-resources.test.ts`；必要的内部 inventory proof 接口只允许由主集成者串行修改 `presentation-project-inventory.ts`。

接口消费真实 inventory、固定 scope/deletionId 和同步 `assertDeleting`；路径只由 project/research/delivery/page/preferences/comments/manual 的固定命名空间生成。重新盘点匹配真实资源身份、独占证明及预算后保存私有文件清单；每次实际 rename/unlink/rmdir 前核原删除 guard、祖先/叶身份。只清自有 quarantine，所有结果经逐资源回执记账。

接受条件：真实合成所有项目独占命名空间被移除；同文档其它项目、文档共享原包、品牌及未知 team/staging 字节不变。中断、持久回执失败及重开不能删错源或重复写未知内容；symlink/篡改/越界拒绝。TDD 先复现缺失能力，再绿色，独立审查后 scoped commit。

## 单元 2：预览、冻结、drain、执行与恢复

文件：`presentation-project-deletion.ts` 和对应 service/integration tests；`presentation-service.ts` 仍由当前唯一 owner 接线。

严格 budget/字段/实际归属先于 admission。预览列出删除项、共享/未知保留项、最小 tombstone/audit 保留策略和确认身份；确认持久同一 intent 后锁外 stop/drain，锁内核原删除版本及归属，逐资源执行和保存回执。失败保留可见 partial；重开按实际 quarantine/source 和现有结果继续。仅全部资源移除或明确解除引用才调用 finishDeletion；不能绕过残留项。

接受条件：两个 Service 实例、不能中断的编译、排队上传/研究/团队、冻结后新 admission、部分失败/重开/重复确认、未知资源出现均有实际测试；不会在持锁时等待依赖该锁退出的任务。完整入口栅栏清单 gate 通过后再公开路由。

## 单元 3：用户确认与保留期执行

文件：实际 Office PC 协议、工作台/runtime 工具与策略服务（接线前明确精确文件）；测试用确定性时钟和临时目录。

用户可见预览/确认/partial/恢复及匿名审计导出消费同一事务。默认 null/null 保留，不默认自动删除；显式内容保留到期使用同一删除执行器，审计保留策略独立处理且不能删最小 tombstone 后复活项目 ID。定时清理只基于真实项目更新时间、已配置策略与记录版本，不依靠测试数量或假时间宣称到期。

接受条件：策略默认保留、精确期限、策略变更竞态、到期任务遇未决/共享资源、匿名审计过期与永久 ID 禁止复活、Office 明确确认、实际 PC 路由多层集成均验证。实机/专业门禁另行验收。

## 迁移、回退与发布

历史资源只通过实际归属核对认领；pre-plan 研究/观察/偏好需要其真实记录加严格控制 scope 证明，不把任意请求 scope 当独占授权。共享引用与 team ACL 需要明确证据；无证据保留并列出。

回退关闭执行能力，保留 intent/quarantine/tombstone/回执，不能清治理记录后允许旧 ID 重写。发布采用分单元独立审查和定向提交，主 factory 接线串行；最终类型、静态、跨层冻结/drain/删除回归及完整相关工作区验证通过后才称工程闭合。整体百分比按原台账，实机未验不计 100%。
