# 单页重做基础：派生修订任务，仅重编目标页

基线de9674c7。依据原方案§5D/E、§6.3、§11.3/11.6、O3、§20/21。原方案不改；高保证/TDD/独立审查与全仓验证。

## 目标/边界

从全部compiled的生产任务创建新修订任务，替换单页SlideIR，保留父任务，其他页PPTX原样继承。该子任务仅目标pending，可调用现有run仅编译该页；失败不丢父成果。复用原计划/样式/资产/主张，标题和claimIds受原计划约束。不是宿主替换、单页重做完整验收或撤销：宿主备份与替换映射下一批接入。

## 接口

production_rebuild_page {operation,documentId,projectId,parentRequestId,requestId,pageId,slide}，requestId与parent不同；slide.id必须pageId。PC读父冻结deck，仅替换该slide，再parsePresentationDeck/assertDeckMatchesPresentationPlan；只允许同一冻结计划，沿用旧revision显示。父必须所有页compiled，否则page_not_ready。无新素材/来源输入；引用父已登记资产/claims，其他页/主题不变。新状态summary沿用schema并可选revision:{parentRequestId,pageId,parentInputDigest}。

Store PresentationProductionRecord新增可选revision上述字段。deriveProduction(projectId,documentId,parentRequestId,requestId,pageId,deck)原子写完整child；验证父全compiled、同id/顺序、非目标slides和顶层非slides内容canonical一致；同child请求/输入/parent/page幂等，其余冲突。复制其他页完整记录与字节，目标pending attempt0；从父继承plan。revision参与严格校验与CAS immutable比较，旧beginProduction不能冒领derived同request，即使deck相同。保留限额32任务/单任务总10MiB/17MiB记录。旧非derived读兼容，无原记录迁移。

Frontend tool rebuild_presentation_page(project_id,parent_request_id,request_id,page_id,slide)请求操作，仅创建，不宿主写入；slide schema取现有PRESENTATION_DECK_SCHEMA.properties.slides.items（验证输入ID/body限额，PC权威parse）。必须校验response revision与输入parent/page，planRevision/summary常规strict。非derived旧summary不带revision；derived任何status/run都保留revision。项目卡显示单页修订来源和目标，提醒尚未替换当前页。单页download可用；prepare遇revision拒绝presentation_page_replacement_required，直到宿主单页替换就绪，防止修订任务被整批重复导入。对旧PC明示upgrade_required。

## 单元

backend负责project-store store+types/tests，TDD原子派生/共享资源变化拒/全compiled限制/parent不变/幂等冲突/旧begin冲突/limits/source-digest验证，scopedcommit。
frontend负责production.ts/project-card及tests，TDD工具schema/严格identity/revision/旧PC/取消/显示警告，scopedcommit。
root负责shell service+helper/runtime注册及refresh、真实8页父→改第2页→只编译1页→解包验证/失败重试父不变/服务重建、plan报告与验收回填。独立复审，npm test/typecheck、变更lint/diffcheck，测试后Addin/Shell构建。保留原PPTX/宿主回执，不部署。新增record字段旧PC可能拒读含修订的项目，发布PC和Taskpane需同步，回滚保留记录不得删成果。
