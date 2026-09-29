# 持久截图记录的本机诊断导出（2026-09-29）

依据原方案阶段0/5及§819诊断默认排除敏感标识；基线8fdb8059。源码`05d5ca53`（5个TS/TSX文件，350行新增、7行删除），Office实际版本`05d5ca53652e`。

## 完成内容

- 用户选择既有“含本机标识”复制诊断时，才惰性读取当前可见任务的持久截图尝试，附独立local_presentation_qa_attempts。记录范围明确为retained_visible_presentation_task，不是全项目/全历史统计。
- 原等待、失败、取消、显式结束和记录完成状态、身份与开始/结束时间完整保留；record_count/unresolved_count仅为保留窗口计数，不生成过去不存在的诊断事件、视觉通过或真实宿主结论。
- 64条/128KiB、版本/状态/规范时间/字段、唯一ID与同一文档/来源/项目/任务/产物摘要严格校验和clone。读取异常/非法记录明确unavailable，不显示虚假空历史、旧成功或原错误文本。无PNG和视觉意见正文。
- 新增本机section超过256KiB总额时保留完整记录与最新会话事件，显式omitted_event_count说明裁剪旧事件数；导出不改变原snapshot或远程事件。
- 普通诊断复制不读取/不含截图账本；远程诊断通道无新增原件数据。无能力/Word/Excel调用形状保持原值。清除当前会话诊断只清volatileevents，下次显式导出仍读当前持久记录。

## 验证

新跨层用例实际先失败：重开后的等待/显式结束/记录完成仍在settings，但诊断没有section；实现后通过。真实PptxGenJS页、实际导入/QA skill、文档binding重开及Diagnostics联调保留原时间和身份，未增加截图/插入或远程发送。UI测试实际createOfficeWorkspaceUi/Diagnostics/clipboard动作，创建和普通复制不读账本，含本机复制才读，异常只显示不可用。

Root专项3文件69/69、A38/38、B27/27及独立最终3文件69/69通过，独立完整5文件审查无正确性/隐私问题，日志`/tmp/c-local-diagnostic-final-review.log`。Root新鲜Office/PC类型、5个变更TS/TSX文件lint/Prettier/diff检查通过。完整相关回归 **3622/3622（302文件）**，79.22s，日志`/tmp/wiswork-screenshot-diagnostics-regression.log`。Office生产构建通过，9.82s，实际版本`05d5ca53652e`，日志`/tmp/wiswork-screenshot-diagnostics-build.log`；常规大chunk提示，未部署。

## 范围、进度与下一步

旧无provider导出的原总量超限错误保持；新增裁剪只用于含本机截图section的导出。测试使用受控宿主/clipboard/PNG，不是PowerPoint真实专业验收。没有部署或上传新增用户记录，旧客户端/混合版本仍待验收。

严格整体64%（575/9，较上轮0个百分点），候选17/20（85%），真实专业0/20。剩余Sony财务原件的三条新官方路径均403，六份183–184字节错误响应未通过PDF文件头。Root独立核对六个错误响应的字节数及SHA，重新运行08部分材料校验通过但仍Sony pending。证据记录于08/acquisition-status.json及/tmp/wiswork-p0-08-acquisition-next/result.json；零份PDF原件，不计候选。已请求可用本机Sony原件和11获授权扫描件，工程任务继续。下一步核对实际诊断错误分类与恢复动作、继续剩余专业材料及原方案实机退出证据。
