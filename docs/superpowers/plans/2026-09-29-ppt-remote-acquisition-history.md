# 网络图片和网页资料获取历史实施计划

## 依据、范围与架构

原方案 §4.3、§4.6、§5.B/E、§7.1、O1/O2 要求真实来源及素材获取状态、失败恢复和重开可见历史。沿用已批准的完整方案和现有本机 PC 持久存储授权。为实际网络图片下载/缓存复用、网页资料下载/解析建立文档级记录；不将获取成功标为事实、许可、品牌或视觉验收通过。不增加宿主写入、自动网络重放、发布或上传。

文档级历史独立保存在 PC userDataPath/presentation-acquisition-history，避免改变旧版附件目录枚举规则。既有附件文档锁保护每次获取的开始/结果，读取历史不等待网络锁。旧 PC 对可选历史读取返回 invalid_request 时，插件隐藏历史能力。保持所有既有获取回执/页任务契约。

## 固定共享接口

- 新子路径 `@wiswork/project-store/presentation-acquisition`：纯浏览器可用 parser 与类型。
- `PresentationAcquisitionHistory`: `{version:1, scope:'remote_material_acquisition', documentId, revision, totalAttempts, records}`。
- `records` 最近 64 条，attempt 按文档全局获取次数连续递增；历史截断由 totalAttempts > records.length 判断。每条 `{id,attempt,kind:'image'|'webpage',source,sourceUrlHash,state:'fetching'|'ready'|'rejected',startedAt}`；source 为去掉查询/片段且无凭据的 HTTP(S) URL，sourceUrlHash 为完整规范 URL 的 SHA256。无原始 URL 查询或原始异常内容。
- ready 额外 `{finishedAt,attachmentId,sha256,sizeBytes,assetSha256?}`；image ready 必须 assetSha256，webpage 不允许。rejected 额外 `{finishedAt,error,attachmentId?}`。error 限 `remote_image_unavailable|remote_webpage_unavailable|quota_exceeded|parse_failed|digest_mismatch|animated_image_unsupported|remote_image_source_conflict|remote_webpage_source_conflict|aborted|invalid_state|acquisition_failed`。fetching 无结束/结果字段。id 为有界安全 ID，时间是规范 ISO；parser exact keys、限制、顺序、互相一致性、深拷贝。
- 新独立 Node Store `PresentationAcquisitionStore`，子路径 `./presentation-acquisition-store`。`read(documentId)` 返回完整历史，`begin(documentId,{kind,source,sourceUrlHash})` 生成随机 ID/起始时间并返回 record；`finish(documentId,id,result)` 原子保存结果，同 ID 同结果读回原结束时间，冲突拒绝。checksum+原子写+256KiB上限，目录和文件防符号链接。
- PC `attachment_acquisition_history` 请求仅 `{operation,documentId}`，响应为完整历史。文档无历史返回空账本；损坏返回既有安全错误，不触碰附件数据。
- 插件 AttachmentSkill `acquisitionHistory(): Promise<PresentationAcquisitionHistory|undefined>`，读取时核对文档，旧 PC invalid_request/upgrade_required 返回 undefined；取消/文档切换/迟到拒绝。HostRuntime/UI 可选方法 `readPresentationAcquisitionHistory` 暴露同一只读入口。

## 实施单元

1. **PC/共享存储**：packages/project-store/src/presentation-acquisition{,-store}.ts 及测试、package exports；apps/shell/src/main/presentation-attachments.ts 获取前/后保存记录、presentation-service.ts 路由及 PC 测试。实际缓存命中也是一次明确的获取请求，ready 表示本机产物通过原有读取验证；动画暂存失败保存 rejected 与附件 ID。网络/解析/容量/取消错误保留原服务回执，同时记录安全失败。发生记录写入异常时保留已发布附件，不重复网络操作。
2. **Office 控制器和 UI**：skills/powerpoint/presentation-attachments.ts、generation.ts/host-runtime.ts、App.tsx 与测试。用户附件区显示可刷新、默认折叠的历史，按 kind+sourceUrlHash 聚合，详情保留每次状态/时间/尝试/结果 ID/安全错误；只有开始记录不推断后台仍执行。失败不清空已上传资料；缺失能力隐藏；不要增加确认门禁。
3. **联合验证与记录**：真实 PC服务→实际Office AttachmentSkill 重开恢复、cache reuse 不重复下载、失败→显式重试、取消、跨文档、64窗口、损坏及旧PC兼容；相关全量、类型、静态、构建、独立审查。源码和验证文档分开提交。

## 回退、发布与验收

不修改已有附件 metadata，旧稿无迁移；回退源码后新历史目录独立保留、旧版本忽略。单元测试先复现缺失行为；不测试实际外网，使用真实服务与注入的响应替代不稳定网络。UI 不进行自动网络重试，历史不构成写入批准。原方案全部实机/专业任务门禁仍需原范围证据，本轮不提高严格进度档位或宣称 100%。
