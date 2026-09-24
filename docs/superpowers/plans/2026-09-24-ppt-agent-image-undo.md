# 图片原图持久备份与确认撤销

基线57fd24ca；原方案§6.2/6.4/O5，既有隔离分支。用户已授权继续实现；实机验收0/20暂缓。本批高保障流程：持久身份、删除边界与中断恢复，TDD、独立审查、全仓验证。

## 目标与架构

在新图片替换前保存原始PNG/JPEG字节到现有PC附件存储；Office设置只存有限元数据。复用既有图片replace/inspectRecovery/finishRecovery做反向替换，恢复原图并验证后才能删除替换图。工作台显示备份及撤销/恢复状态，所有宿主写入经确认。无需新依赖或新存储框架。

范围为已绑定生成/生产页中的受支持普通原生图片，沿用原有2MiB媒体、32条/128KiB设置预算和附件配额。不承诺裁剪/效果/链接图片支持、旧记录补造备份、原shapeId恢复、跨类型多级历史或任意现稿基线。

## A：图片保存点与反向恢复

文件：powerpoint-package.ts、browser-presentation-image-adapter.ts、presentation-image-replacement-record.ts、presentation-document.ts、presentation-page-editing.ts及对应测试。

新增adapter.captureOriginal(slideId,shapeId,signal):Promise<{snapshot:PictureSnapshot;base64:string}>，从同一已核验单页导出提取原始媒体；inspect不附加媒体字节。备份与before须完全匹配，写前重新核验。

新增options.imageBackup接口available():boolean；save(documentId,base64,signal):Promise<{attachmentId:string;sizeBytes:number;mime:'image/png'|'image/jpeg'}>；load(documentId,backup,signal):Promise<string>。配置存在但不可用时禁止新图片替换；运行时总配置此接口。保留未配置接口的既有低层调用兼容性，不对其承诺撤销。

记录扩展可选backup、after（替换后PictureSnapshot）、undoBaseline、restoredShapeId，state增加undo_pending/undone。严格字段/状态/身份验证，旧记录可读不可撤销；不可变内容、队列保存/SaveAs/预算继续适用，预留反向记录空间。after须在原替换完成与恢复完成时从宿主读回，防止撤销覆盖手工变动。新工具undo_presentation_image_replacement，复用inspect/resume处理undo_pending，输入project_id?,page_id,shape_id（历史oldShapeId）,explanation?。complete有完整证据可undo；undo_pending inspect/resume；undone只历史。不自动重插未知ID。

撤销核验备份摘要、当前after完整快照及绑定→保存undo_pending和undoBaseline→使用既有replace插回原图→持久restoredShapeId→删除替换图→保存undone。反向恢复构造以newShapeId为old、restoredShapeId为new、原mediaDigest为assetDigest的pending记录复用adapter恢复；无restoredShapeId只人工检查。取消、原图保存失败、备份丢失/损坏、写后回执失败、手工改动均不得跳过核验或自动重插。

TDD：原始字节/不支持图；替换前保存失败零宿主写；撤销重开/确认/冲突/反向恢复/只补回执；记录预算/CAS与旧版兼容。限定提交。

## B：原始附件读取与备份客户端

文件：shell presentation-attachments.ts、presentation-service.ts与必要操作路由；office新增presentation-image-backup.ts；各自测试。

复用attachment_begin/chunk/finish存储（确定性文件名image-backup-<digest>.png/jpg；同digest已有附件复用其metadata文件名），新增attachment_original分块读取{attachmentId,offset,length<=128KiB}，仅ready image，严格document作用域、no-follow、完整原始摘要/大小核验，不读取标准化image.png作为原图。返回{attachmentId,offset,sizeBytes,sha256,mime,base64}。客户端createPresentationImageBackup({available,request})返回A接口；2MiB限制，规范base64，内容摘要与返回元数据一致，每步检查取消；重开后load不依赖VFS，拒绝缺失/过期/错误返回，不输出base64到Agent文本。每次save完成后原始读回核验才返回。

TDD：真实服务存储重启后读回原始字节，标准化产物不同仍返回原图，跨文档/缺失/损坏/配额/越界/符号链接/中断及客户端坏响应。限定提交。

## Root：接线、工作台、跨层

host-runtime.ts配置imageBackup（附件availability/request），新工具路由及QA失效；presentation-changes.ts/card支持完整记录undo，undo_pending检查/恢复，undone正确显示恢复对象/原摘要。旧complete不可undo，状态旧响应/切换保护继续生效。实际runtime+binding+提案+备份服务跨层覆盖替换→重开→撤销及回执恢复，不仅mock工具分派。

## 验证与交付

独立审查完整diff并修复重要问题；全仓test/typecheck/lint/licenses/format/diff及Office/Shell构建。原方案不变，更新阶段报告/审计/进度台账，以原验收口径评分，不靠测试数抬高进度。旧终态可读，旧pending继续既有恢复；新undo_pending旧版本无法识别，降级不能声称自动恢复。备份保留，回退代码不删用户数据。仅隔离分支提交，不合并/推送/部署。
