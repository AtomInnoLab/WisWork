# 单页替换：可恢复暂存与撤回

基线e0e65046。遵循原方案§6.4/O3/O5/§21，继续补写入事务和可撤销要求；原方案不改。本批是单页替换的暂存阶段，原页始终保留，不更新业务页映射或宣称完成替换。既有隔离分支，高保证/TDD/独立复审/全仓验证。

## 目标与设计

根据明确的ready原页备份和已编译子修订，在原页之后插入该修订单页。先持久化pending，再写入新页，唯一新页ID回读后先持久化inserted，验证包内容后staged。保留原页和原回执；明确确认可撤回暂存页。中断后只读检查；已知inserted可重新验证后完成staged，不再插入；pending不自动重试。discard_pending可在新页仍匹配时继续删除或新页已消失且原序已恢复时仅补回执。

包内容指纹：复用powerpoint-package的受限ZIP读取，对所有非目录项按路径排序后计算路径+条目SHA256的SHA256；忽略ZIP容器时间戳/压缩差异，不忽略OOXML内容、关系或资源变化。与Office实际导出不一致则停留不确定状态，不删除原页。真实Office规范化兼容待实机验证。

## 持久记录

PresentationPageReplacement {version:1,changeId,documentId,projectId,parentRequestId,requestId,pageId,backupId,parentArtifactDigest,backupDigest,originalPackageDigest,replacementPackageDigest,sourceSlideId,oldSlideId,beforeSlideIds,state:'pending'|'inserted'|'staged'|'discard_pending'|'discarded',newSlideId?}。

严格键/ID/摘要/页序/状态校验，最多512原页且new不在before内，sourceSlideId合法。单记录<=192KiB，为新ID预留空间。document settings独立key，CAS串行保存，身份不可变；pending→inserted（首次newID）→staged→discard_pending→discarded，同值幂等。只有discarded后允许新change；未决记录跨项目阻断。旧记录无迁移。备份和生产成果不删除。

## 单元及接口

1. journal：新增presentation-page-replacement-record.ts和binding测试；修改presentation-document.ts返回readPageReplacement()/writePageReplacement(record,expected)，沿用已有receiptQueue、document/SaveAs检查与失败回滚。独立限定提交。
2. adapter：新增browser-presentation-page-replacement-adapter.ts及测试；powerpoint-package.ts新增presentationPackageDigest(base64,signal)。BrowserPresentationPageReplacementAdapter.inspect(record,signal)返回{status:'baseline'|'staged'|'conflict',slideIds}：baseline匹配原序及原包，staged匹配原序中old后唯一new且两包指纹正确，否则conflict。stage(record,base64,onInserted,assertCurrent,signal)返回void：插入前重新验证原内容/顺序及文档，单页KeepSourceFormatting插old后，回读唯一增量，调用onInserted(newId)保存后验证内容。discard(record,assertCurrent,signal)删除已匹配暂存页并验证恢复原序；若baseline已恢复则不写入。任何不确定状态保留日志，不自动清理其它页。
3. tools：新增presentation-page-replacement.ts及测试。options:{available,request,documentId,artifact,readReceipt,readPageReplacement,writePageReplacement,loadBackup,adapter,proposals}。stage_presentation_page_replacement输入project_id/request_id/page_id/backup_id/change_id；inspect/resume/discard_presentation_page_replacement输入project_id/change_id。stage校验完整父导入回执及摘要、备份绑定和字节、编译子revision与目标页，提出明确预览；确认时重验后写pending、调用stage、写staged。后续工具以持久记录+当前父artifact/receipt为绑定，resume只允许inserted且inspect为staged，discard只允许staged/discard_pending；均确认，CAS和epoch取消贯穿。pending/conflict保持人工核对。所有写入局部QA失效，工具不得启用整批派生导入。

root负责backup.loadBackup复用下载逻辑（不落VFS），导出production parsePageArtifact供校验复用、runtime注册/生命周期/局部QA、实际documentbinding wiring、集成测试/报告。

## 验证与发布

TDD包括已存pending才插入、inserted保存失败不重复插入、内容/页序/文档漂移阻断、成功暂存与撤回、discard_pending恢复仅补回执、清会话使旧确认失效。受限ZIP指纹覆盖容器变化相同/内容变化不同。全仓npm test/typecheck/lint、format/check、licenses，再构建Addin和Shell；独立复审。模拟Office不能替代实机，20项专业任务与P0退出未完成。

不部署/合并。新settings key只由新插件使用，旧版本不应操作未决事务；回滚保留日志和PC备份，暂存原页始终存在。最终删除原页、映射切换、完整整页撤销和QA需后续单元接入后才开放正式替换。
