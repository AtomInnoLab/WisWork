# 单页重做：写前原页持久备份

基线fffb58fc，沿用已批准原方案§6.3/6.4、O3、O5和§21。高保证：既有隔离分支、TDD、独立复审、全仓验证。

## 目标与边界

将真实宿主页的完整单页PPTX保存到PC，绑定已编译单页修订、父任务、业务页、宿主页和页序；中断续传，重启后可读取下载原始字节。它是历史保存点，不证明页面当前仍未改变。不进行页面替换、删除、整页撤销或QA通过；保留派生任务整批导入阻断。后续替换必须重新核对当前原页内容与位置，再使用备份和事务回执。

## 架构与约束

Taskpane用明确的修订请求和当前准备好的父任务导入回执解析稳定宿主页ID。Browser adapter在一次Office运行中按稳定ID导出并核对页序。PC保存不可变元数据与原始字节，通过128KiB分块传输；完成时校验SHA-256和单页PPTX结构，不重写包。全部操作绑定documentId/projectId/backupId及已编译修订父子关系。每备份8MiB，每项目8个备份（包括未完成）；最多512个宿主页ID。沿用256KiB请求限额。无自动清理覆盖，无秘密数据日志。

接口：page_backup_begin携带backupId/requestId/pageId/hostSlideId/slideIds/sha256/sizeBytes；page_backup_chunk携带backupId/offset/base64；page_backup_finish、page_backup_status携带backupId；page_backup_read携带backupId/offset/length。所有操作均带operation/documentId/projectId。响应元数据包含上述绑定与parentRequestId/parentInputDigest/inputDigest/status/receivedBytes；read返回backupId/offset/sizeBytes/sha256/base64。禁止把调用方声明当成Office真实性证明。

## 可审查单元

1. PC：新增apps/shell/src/main/presentation-page-backups.ts及tests/presentation-page-backups.test.ts；严格协议、生产记录绑定、不可变元数据、原子完成、校验损坏/大小/数量/路径/符号链接、分块重放、重启。先失败测试后实现，独立限定提交。root接service分派。
2. Taskpane：新增skills/powerpoint/presentation-page-backup.ts及tests/presentation-page-backup.test.ts；save_presentation_page_backup与read_presentation_page_backup，父导入绑定校验，读取子status，稳定ID导出、上传/续传、下载哈希验证至VFS；不暴露base64到模型。生命周期epoch和document/artifact/receipt重验。先失败后实现，限定提交。
3. root：BrowserPowerPointAdapter.exportPresentationPagePackage(slideId,signal)返回{slideId,slideIds,base64}，稳定身份和页序检查；host-runtime注册/clear；service注册；跨层/adapter验证、报告与方案对应表更新。

## 验证与发布

定向测试覆盖真实单页PPTX、重启与断点、身份/内容冲突、损坏输入、用户切文档、清会话，不依赖真实Office。独立复审完整差异，npm test、npm run typecheck、变更ESLint、git diff --check，随后构建Office Addin和Shell。实际Office兼容及保存重开仍待实机验收，不宣称P0完成。

PC新增独立backup目录，不迁移旧记录；旧PC拒绝新操作；显式unsupported映射升级提示，通用invalid_request保留错误。Taskpane与PC后续同步发布，当前不部署/合并。回滚只停用新工具，保存所有备份，不删除恢复资料。PC本地磁盘丢失不在本批备份保证内。
