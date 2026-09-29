# 当前 PowerPoint 文稿与完整交付包

依据已批准原方案 §5.G、§7.1、§8.5、§11.8、O1/O4，沿用隔离实现分支；不更改原方案。旧导出来自编译快照，不能代替当前宿主文稿。本批建立真实宿主 PPTX、可选宿主 PDF、证据、主张、来源、历史验收和保存点的可恢复交付包；未证明关键门禁时明确为“检查待完成”，不生成 project.completed。

## 固定接口

1. `apps/office-addin/src/skills/powerpoint/presentation-document-export.ts`：`supportsPowerPointDocumentExport(format:'pptx'|'pdf'='pptx'):boolean`、`exportPowerPointDocument(format:'pptx'|'pdf',signal?:AbortSignal):Promise<Uint8Array>`。实际 Common API getFileAsync(Compressed/Pdf)，64KiB切片，PPTX最多20MiB/PDF最多10MiB；检查声明长度、切片序号/尺寸/字节、PPTX ZIP或PDF签名；有界回调等待，所有成功取得的File都closeAsync，包括取消后迟到对象。原稿无写入，不调用编译器。
2. 纯共享契约子路径 `@wiswork/project-store/presentation-delivery-bundle`，parser/type：`PresentationDeliveryBundleManifest` = `{version:1,scope:'current_office_document',documentId,projectId,requestId,planRevision,inputDigest,planDigest,createdAt,files,checks}`；files每项 `{name,sizeBytes,sha256}`，mandatory `presentation.pptx,evidence.json,evidence.md,claims.json,sources.json,quality.json,checkpoints.json,README.md`，optional `presentation.pdf`，不包含manifest自身（ZIP另有manifest.json）。checks固定 `{completion:'not_verified',sourceAuthority:'not_verified',timeliness:'not_verified',roundTrip:'not_run',hostQa:'not_checked'|'historical_records_only',pdf:'included'|'not_requested'|'unavailable'}`。所有未知/多余字段、重复文件名、非法时间/摘要/身份拒绝；metadata不作为宿主真实性/完成证明。
3. `PresentationDeliveryBundleReceipt` = `{version:1,documentId,projectId,requestId,bundleId,sha256,sizeBytes,receivedBytes,state:'uploading'|'ready',createdAt,completedAt?,manifest}`；bundleId===sha256（ZIP SHA256），zip最多20MiB，未压缩文件合计最多32MiB；ready receivedBytes===sizeBytes且完成时间规范。parser名 `parsePresentationDeliveryBundleManifest`, `parsePresentationDeliveryBundleReceipt`。
4. PC独立本机缓存 `userDataPath/presentation-delivery-bundles/{hash(doc)}/{hash(project)}/{hash(bundleId)}`；操作前缀 `delivery_bundle_`，通过现有presentation.v1，不新增Relay协议/能力。operations：
   - `begin` body `{operation,documentId,projectId,requestId,bundleId,sha256,sizeBytes,manifest}` -> receipt。
   - `chunk` body `{operation,documentId,projectId,requestId,bundleId,offset,base64}` -> receipt；128KiB块，有界幂等/overlap前缀核对，fsync后ACK。
   - `finish` body基本身份 -> ready receipt；核对总长度/SHA/ZIP固定文件集、各摘要/大小、manifest与begin一致。校验 evidence.json对应真实冻结production身份、计划和页列表。只发布完整包。
   - `metadata` -> receipt；`read` adds offset/length -> `{bundleId,offset,totalBytes,base64}`（最大128KiB）；`list` `{operation,documentId,projectId,requestId}` -> `{bundles: receipt[]}`最多32条；`delete`基本身份 -> `{bundleId,deleted:true}`，供用户明确清理本机包，原PPT不变。
   - 32个包/项目、100MiB预留quota/项目（未完成包也预留）；重复begin相同内容幂等，冲突拒绝；文件/目录无符号链接、原子metadata+校验、有界Zip预检查/解压、防路径和重复entries；损坏不覆盖。
   - 项目status planned/compiled返回 `deliveryBundlesAvailable:true`；旧PC absence隐藏功能。服务独立handler在项目常规操作前严格路由，ProjectStore绑定doc/project/production，未修改原输入或编译页。
5. 根代理实现 `createPresentationHostBundleSkill(options)`（新 `presentation-host-bundle.ts`），options `{available,request,documentId,vfs,nativeAvailable,exportDocument,readQuality?,readCheckpoints?}`。readQuality/projectId/requestId只读取当前文档已有 QA 记录；readCheckpoints只读取现有变更历史；均是历史范围，不伪造当前宿主或round-trip通过。tools `export_current_presentation_bundle`（project_id,request_id,include_pdf:boolean可选默认false）与 `restore_presentation_delivery_bundle`（project_id,request_id,bundle_id）；clear取消等待。ZIP固定manifest+files，真实当前宿主PPTX；PDF实际宿主不可用时包仍保留PPTX且checks.pdf=unavailable，禁止用编译PDF替代。文件/ZIP摘要绑定，上传PC后原子写会话VFS ZIP，响应丢失提示刷新/读取，不自动重放宿主导出。
6. UI/Runtime：PresentationGenerationOptions新增 `deliveryBundlesAvailable?():boolean`（App记录当前project status?可用PC status标记，Runtime可先readonly probe `status`项目再提供动作）；实现时可直接ProjectController status有字段才显示动作，Skill内部校验status能力。HostRuntime注册新Skill，浏览器native helper，历史QA回调及checkpoints读取；ProjectController新增可选 `exportCurrentBundle?(includePdf?:boolean)`、`restoreDeliveryBundle?(bundleId:string)`、`readDeliveryBundles?()`，snapshot `deliveryBundles?:receipt[]`。工作台显示当前宿主交付包与可选宿主PDF、只读刷新/重开恢复及从PC恢复ZIP到会话附件。动作可用以status标记与native helper为准；旧PC不显示。删除本机包需用户明确确认后调用PCdelete（可本批先提供skill/helper，UI不要自动清理）。错误中文，phase pending/cancel/文档切换/迟到保护。原来编译PDF入口保留并明确其来源。

## 实施与验证单元

A. PC/shared：共享parser、缓存服务和delivery_bundle操作、status能力字段、package exports、存储/服务测试。RED后实现，测试完整 ZIP各文件校验、续传/重复/损坏/错误身份/容量/清理及冻结计划匹配。禁止替native宿主生产文件。
B. native helper：实际getFileAsync/getSliceAsync/closeAsync适配器，SDK回调夹具覆盖真实切片、取消、迟到、超限/异常/unsupported及资源释放；查官方Common API文档，不修改Word导出。
C. Runtime/UI：新Skill注册、历史QA/保存点回调、project controller及实际工作台动作/历史恢复。消费根代理提供固定Skill接口，不改PC/native helper。先RED，真实React及controller验证旧PC、文档/取消、响应丢失、重开恢复。
D. 根代理 orchestrator与真实跨层：native文件字节→JSZip交付包→真实PC缓存→新Taskpane从PC恢复；修改后的host内容保留，所有文件摘要一致，关键门禁未验证不生成完成。执行全相关回归、类型、静态、构建和独立审查，源码与进度文档分别提交。

本机数据由用户明确导出操作经已配对PC持久保存，资料不发送新的外部服务。测试使用合成PowerPoint字节及本机临时目录。20MiB包/会话单文件限制超出时明确报容量，不截断或替换文件。回退源码时独立缓存保留，旧版本忽略；没有上线、上传或部署操作。实机/专业20任务仍按原标准验收。
