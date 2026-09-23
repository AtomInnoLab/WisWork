# O3：页生产成果的宿主导入与恢复

基线 f65f21db。依据已批准完整方案§4.7、§5E、§8.7、§11.6、O3、§14、§20/21。复用现有Office页导入事务，不合并OOXML，不改写原方案。

## 目标与架构

将完成编译的页任务准备为有序单页PPTX集合，复用一次确认、逐页预写回执、插入、ID回读及不确定写入停止的机制。独立工具与回执命名空间防止与旧整稿同名请求混淆。业务pageId绑定页索引，即使各文件sourceSlideId相同也能恢复。

本批仅接完整已编译任务到追加式导入；不实现边编译边写入、替换页、内容/渲染QA、保存重开、撤销。原有整稿QA/修改工具继续使用其原成果，页生产QA映射将在后续接入，不能冒称完成验收。

## 交付单元

1. delivery与持久化（backend）：扩展CompiledPresentationArtifact可选pagePptxBase64:string[]，只有生产集合使用，pptxBase64为空。checkpoint v2新增pageIds:string[]，sourceSlideIds允许重复；v1保持原校验。新工具import_presentation_production/read_presentation_production_import_status复用现有循环；从数组取对应PPTX和真实sourceSlideId，完成映射使用稳定pageIds。digest绑定完整集合和元数据；回执key production/project/request。document imports严格识别v2与命名空间，兼容已有v1。文件presentation-delivery.ts、presentation-page-delivery.ts、presentation-document.ts及对应测试。测试重复源ID、多页resume、不确定不重写、命名空间隔离、改字节/顺序/业务ID拒绝、保存重建；先RED再GREEN，范围提交。
2.准备（frontend）：presentation-production.ts与其测试，新增prepare_presentation_production_import(project_id,request_id)，先读严格status，全部compiled后顺序拉每页，复用现有响应校验；planRevision与业务pageId一致、累计原始字节<=10MiB；每await检查身份/epoch/abort，完成并remember后原子发布有界缓存，不触碰generation artifact。返回artifact(projectId?)，clear清缓存；失败不能发布半成品。准备重跑可恢复旧任务；工具不写宿主、不标QA。先RED再GREEN，范围提交。
3.root：host-runtime注册/路由production import，与原整稿独立；importProgress卡展示最近显式准备的产物（旧生成操作切回旧成果），clear取消选中。真实PptxGenJS→PC→准备→宿主adapter测试，检查一份文件一次选页、恢复、ID映射及不会混同旧成果。阶段报告回填验收。

## 验证与发布

独立复审完整差异；全仓npm test、npm run typecheck、变更ESLint与diffcheck，测试结束后Addin/Shell构建。Office适配用mock验证，不称实机通过；真实20任务仍待执行。无PC/Relay新协议，无旧数据迁移，新增v2回执旧版会拒绝，部署/回滚必须保留回执且避免旧客户端恢复生产导入；未合并/发布。报告明确兼容限制和下一阶段入口。
