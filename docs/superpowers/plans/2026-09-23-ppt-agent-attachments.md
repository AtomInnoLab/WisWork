# PPT Agent：文档附件可靠上传与解析

基线：134154c；沿用已批准的完整方案，推进 O2 / 阶段 2 的文档附件链路。

## 本轮闭环

Taskpane 文件选择 → 128 KiB 顺序分块 → PC 文档隔离持久化 → 复用 file-parse → Agent 分页读取文本 → 附件 URI 可写入来源台账。支持 PDF、DOCX、TXT、MD、CSV、JSON；图片素材、网页抓取和 OCR 留待后续。

附件在计划建立之前即可上传，因此按持久化 PowerPoint documentId 隔离，不依赖 projectId。来源 ID 为文件 SHA-256，URI 为 attachment:<hash>，解析结果不代表核验通过。服务只接受标识和字节，不接受本地路径或远程 URL。

## 协议与边界

复用现有 JSON 传输，独立协商 presentation-attachments.v1 能力（旧 PC/Relay 不具备时保留会话附件路径），添加 attachment_begin/chunk/finish/list/read 操作。每次请求带 documentId。begin 固定 name/sizeBytes/sha256/attachmentId；chunk 带 offset/base64，重复块应幂等；finish 校验整体摘要并缓存解析结果；list 返回当前文档附件；read 返回 offset/maxChars 限制的纯文本。

单文件 50 MiB，单块 128 KiB；每文档最多 32 件、声明总量 100 MiB；文本最多 100 万字符，单次读取最多 24000 字符。原始内容、上传进度和解析结果持久化，重启后重选同文件可续传。会话清空取消本地操作，不删除 PC 的持久附件。

## 实施单元

1. 后端独立单元：持久化、幂等与摘要、目录隔离、解析缓存、配额、损坏与取消测试；子代理实现并提交。
2. 前端独立单元：上传客户端、Agent list/read、现有上传入口、生命周期与响应校验；子代理实现并提交。
3. 集成：主服务路由与错误边界、真实 PDF/DOCX 通路测试、依赖声明、阶段报告。
4. 独立审查完整变更，修复重要发现；全仓测试与类型检查、变更文件 lint、插件和 PC 构建依次执行。

## 验收与限制

验证原始字节经分块持久化后真实解析、PC 服务重建后仍可读取、错误文档不可读取、重复块/丢失响应可恢复、失效会话不发布附件。保留已有编译/计划/导入行为。解析器复用不等同沙箱隔离；记录输入与解压边界，以及尚未提供的 OCR/视觉 QA。工作留在隔离分支，不自动合并或发布。

## 执行证据

- 后端独立单元提交 `8548c04`；前端独立单元提交 `93f828c`。
- RED：新增真实附件服务调用因 unsupported operation 返回 invalid_request；GREEN：真实 PDF、Word、50 MiB 边界与断线恢复三项集成测试通过。
- 独立审查发现旧 PC 能力误判；新增独立附件能力协商，混合版本回归测试通过，复审无剩余重要问题。
- 全仓 5751 项 Vitest 测试、Relay 27 项 Rust 测试、全仓类型检查、变更文件 lint、Rust 格式、两个生产构建均通过。
- 分支保留在现有隔离工作区；未合并、部署或进行真实 PowerPoint 宿主验收。详细限制和下一步见 product 目录附件实施小结。
