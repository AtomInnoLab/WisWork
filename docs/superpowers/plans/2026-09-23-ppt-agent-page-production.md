# O1/O3：持久化页级编译任务

基线f73f02f7。依据批准完整方案§4.5、§5阶段E、§11.3/11.6、O1/O3、总体阶段1/3、§20/21；原方案为验收权威。本批完成PC页任务与Taskpane状态，不宣称完整连续交付、宿主单页替换或P0验收。复用已批准设计，不重复范围审批。

## 设计与范围

新页生产任务冻结已保存计划与整稿SlideIR，顺序调用现有PptxGenJS编译器，每次仅编译一页及其引用资产/主张。成功页持久化单页PPTX；失败页记录错误，后续页继续。重试相同任务只编译未完成页；进程中断遗留building可重新编译（无宿主写入）。不合并单页OOXML：现有mergeSlideFromPptx会丢notes和非media关系，不符合图表/来源保真要求。

新文件与旧整稿编译回执隔离，不改变旧下载、导入映射和用户文稿。单页编译不代表导入/QA/交付；下一批将页成果接入逐页交付，保留原方案中的单页重做/撤销等缺口。

## Store接口（backend owns）

在packages/project-store/src/presentation-store.ts与index.ts新增类型及方法，复用现有私有bind/read/write/canonical/hash；不新建通用框架。

PresentationProductionPage={pageId:string,state:'pending'|'building'|'compiled'|'failed',attempt:number,error?:string,result?:{pptxBase64:string,sourceSlideId:string,report:unknown},resultDigest?:string}。
PresentationProductionRecord={version:1,projectId,documentId,requestId,sequence,inputDigest,deck:unknown,plan:PresentationPlanBinding,planDigest,pages:PresentationProductionPage[]}。

beginProduction(projectId,documentId,requestId,deck,planBinding):record；相同请求和输入/计划幂等，冲突拒绝。production(projectId,documentId,requestId?):record|undefined，不传请求取sequence最新。
updateProductionPage(record,pageId,update:Pick<PresentationProductionPage,'state'|'attempt'|'error'|'result'>):record；CAS比较最新存储页与record中的原页；其他页用当前存储保留。pending/failed/building→building且attempt+1；building→failed/compiled且attempt不变；compiled不可改。error限枚举'compile_failed'|'invalid_deck'|'aborted'|'output_too_large'|'asset_unavailable'。building/compiled不带error，failed必须error，compiled必须result及内部计算SHA256；其他状态不带result。缺页/错误序列/摘要/重复ID拒绝。

只接收包含1..32唯一slide.id的deck；身份仍绑定文档。单记录<=17MiB，累计compiled PPTX原始字节<=10MiB，创建最多32个生产请求。严格读取、canonical摘要、symlink拒绝、atomic写入。结果sourceSlideId合法numeric#，base64严格且有界，report有界JSON（可unknown但不允许undefined/非法JSON）；读时校验digest。文件production-<sha256(requestId)>.json，旧读路径忽略。只暴露有界数据，冻结plan/deck不可写回。

## PC操作/响应（root owns）

在presentation-service.ts复用同project锁与attachments，接入辅助模块presentation-production.ts：
production_begin {operation,documentId,projectId,requestId,planRevision,deck} 必须匹配当前已保存计划；已有同request重试使用当时快照（与旧整稿语义一致）。
production_status {operation,documentId,projectId,requestId?}，无request取最新。
production_run {operation,documentId,projectId,requestId} 顺序跑剩余页；单页失败不阻后续；取消停止后续且不返回虚假compiled。
production_page {operation,documentId,projectId,requestId,pageId} 仅返回成功页。
前三项返回summary={projectId,requestId,planRevision,status:'pending'|'building'|'partial'|'compiled',compiledCount:number,total:number,pages:[{id,title,state:'pending'|'building'|'compiled'|'failed',attempt:number,error?:枚举}]}。全compiled→compiled，有building→building，有failed或部分compiled→partial，否则pending。
page响应={projectId,requestId,pageId,planRevision,status:'compiled',pptxBase64,sourceSlideId,report}。
旧status操作增加可选production:summary（包括只有plan没有整稿请求的项目）。旧客户端忽略附加字段；旧PC收到新op明确unsupported/invalid_request，不伪造支持。不新增协议名称或放宽传输上限。

资产仅解析本页引用，复用PC文档隔离缓存，单页及总成果有界。compile结果需1页sourceID且报告slideCount=1/deckId匹配、bytes非空<=10MiB；缓存返回做完整校验。请求仍<=256KiB、响应<=15MiB、无主进程并行编译。

## 插件（frontend owns）

新增presentation-production.ts技能（options用PresentationGenerationOptions，clear epoch）；tools：start_presentation_production(request_id,deck,plan_revision必填)、run_presentation_production(project_id,request_id)、read_presentation_production(project_id,request_id?)、read_presentation_page_artifact(project_id,request_id,page_id)。映射上述4op，strictinput、响应逐项validate与请求identity一致；每次await后核验epoch/doc/available；rememberProject后再发布。有素材引用的begin需assetsAvailable。read artifact写专属bounded VFS文件，返回path和编译报告，不写generation.artifact映射。无需Office写确认（仅PC准备/文件读取），明确尚未导入/QA。

扩展presentation-project.ts status可选production，复用新技能模块export parsePresentationProductionStatus(value)；project-card显示每页待制作/制作中/失败待重试/已编译（未导入验收）、完成数、刷新与继续页任务按钮。controller新增runProduction(requestId)使用execute('run_presentation_production'...)，不混同resume整稿。运行期间已有cancel生效。

root owns host-runtime.ts 注册/路由，clear，生产工具完成后refresh项目卡。frontend不改host-runtime/App（现有project card已嵌入无需新入口）。

## 单元/验证

1.backend Store+types+tests，TDD：幂等/跨文档/plan冻结/CAS/每页state/缓存digest/限额/服务重建可读；scopedcommit。
2.frontend skill+project controller/card+对应tests，TDD：严格协议/身份/取消/迟到/部分失败/恢复/旧PC错误/单页文件下载；scopedcommit。
3.root PC service/helper+跨层真实8页测试/失败注入重建只补失败页、runtime、阶段报告；scopedcommit。

独立复审、全仓test/typecheck、全部变更lint、测试完成后Addin/Shell构建。工程测试不替代真实PowerPoint验收。未合并/部署；新增独立文件旧版忽略，回滚保留成果。阶段报告回填P0清单并保留未完成项。
