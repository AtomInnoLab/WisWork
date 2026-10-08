# O1/O4：当前 PowerPoint 文稿与可恢复完整交付包

日期：2026-09-29。依据原方案 §5.G、§7.1、§8.5、§11.8 和 O1/O4；原方案及退出门槛保持不变。

## 完成范围

1. 新增 PowerPoint 原生 Common API 全文稿导出适配器，真实调用 `getFileAsync(Compressed/Pdf)`、64KiB 切片和 `closeAsync`，导出当前宿主文稿而非编译快照。校验声明大小、切片序号/字节/长度、累计大小和格式签名；PPTX 20MiB、PDF 10MiB 上限，有界等待、取消、超时和迟到 File 释放。畸形 SDK 回调、非 Error 故障与重复文件回调有回归覆盖。未修改 Word 导出或原文稿。
2. 完整 ZIP 包包含当前 `presentation.pptx`、可选原生 `presentation.pdf`、`evidence.json/.md`、`claims.json`、`sources.json`、`quality.json`、`checkpoints.json`、`README.md` 和独立 `manifest.json`。逐文件与 ZIP 的 SHA-256 绑定冻结生产身份、计划修订和摘要；JSON/Markdown 使用既有完整交付报告。质量记录严格限定同文档、同项目、同任务；保存点保留实际历史元数据和本机备份引用。
3. 整份宿主文稿可能含本项目以外页面或用户修改，冻结计划证据可能只覆盖部分页面。历史 QA 不提升为当前宿主 QA；保存点包只含历史元数据、不含备份文件，不是独立可还原备份。PPTX/PDF 分别读取，导出期间修改可能使两份快照不同，包内明确一致性尚未验收。所有关键检查保持 `not_verified` / `not_run`，不生成 `project.completed`。
4. 新增独立本机 PC 交付包缓存，通过现有 `presentation.v1` 实现 begin、chunk、finish、metadata、read、list、delete。128KiB 块、重叠核对、幂等上传、fsync 后确认、校验和元数据与原子发布；重开或响应丢失后可只读查询，并恢复完整 ZIP 到新会话附件，不自动重复原生导出。
5. ZIP 在解压前核对中央目录、本地头、重复文件、路径、ZIP64、加密和膨胀上限，随后核对 CRC、固定文件集合、各文件摘要、完整 manifest、真实冻结 production 计划/页身份及 claims/sources。符号链接和非普通文件拒绝。包最大 20MiB、未压缩合计最大 32MiB；每项目最多 32 包、预留总量 100MiB，未完成上传也计入。
6. 原子临时文件及 staging 受项目锁和严格命名约束；重开清理崩溃遗留临时文件，不清理有效成果、任意文件或其他项目进行中的上传。工作台提供删除本机包入口，必须用户明确确认，支持清理未完成上传；不删除原始 PowerPoint 或会话附件。
7. 实际 HostRuntime、Agent 工具注册、ProjectController 和 React 工作台接通导出、可选 PDF、只读刷新、重开列表与恢复 ZIP；严格文档/项目/任务隔离，取消和迟到守卫。旧 PC 不声明能力时隐藏。原生 PDF 不可用仍保存 PPTX，保留明确提示，禁止替换为编译预览 PDF；原编译 PDF 入口继续标明来源。配额、文件容量、宿主支持和历史异常以安全中文提示展示。ACP 活动标记文件交付操作，不把操作完成宣称为项目完成。

官方 SDK 依据：[PowerPoint/Word 全文稿导出](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word)、[Office.File](https://learn.microsoft.com/en-us/javascript/api/office/office.file?view=common-js) 与 [Office.Slice](https://learn.microsoft.com/en-us/javascript/api/office/office.slice?view=common-js)。宿主支持按实际 requirement set 与回调验证；当前 Linux 测试使用真实 SDK 形状的合成回调，不是实机 PowerPoint 验收。

## 验证

失败先行覆盖新增共享契约、PC 服务、原生导出、工作台能力及删除入口。定向验证：原生 SDK 适配 **24/24**、PC 缓存 **11/11**、包编排 **9/9**、实际 Runtime/Controller/React **124/124**。两项真实跨层测试连接原生 SDK 形状回调、实际 Office Skill、真实 PC 服务与持久存储，验证用户修改后的字节完整保留、两端重开恢复、保存后响应丢失只读找回，且未调用编译器。

完整相关回归首次发现两条旧接口枚举断言：原工作台测试不允许新的只读 `delivery_bundle_list`，原 PC 状态精确对象没有新能力字段。保留原行为检查并补上新接口后，最终 **2661/2661（198 文件，95.31 秒）**全部通过。三个相关包 TypeScript、所有改动 TS 的 ESLint、格式和差异检查通过。

独立 broad review 未发现严重或重要问题，提出崩溃临时文件残留的轻微问题；补反例、项目锁下严格清理及符号链接拒绝后，增量复审确认问题解决。新增明确确认删除与分别读取 PDF 快照的说明也通过独立复核，无遗留审查问题。

源码提交 `200098c6`，共 21 个源码/测试/契约文件。Office 生产构建通过（9.21 秒），产物版本对应源码 `200098c6c6a1`；构建保留既有 chunk 大小提示，未部署。

## 进度与下一步

总体 **64%（575/9，较上轮 0 个百分点）**，O1/O4 **75%**，候选材料 **14/20**，真实专业任务 **0/20**。本批完成原方案交付工程闭环，不代表全部需求或原退出验收完成。

下一步继续完整研究结果与语义事件绑定、关键内容/来源/时效核验及当前宿主与保存重开检查；落实原 20 项真实专业任务与各模块退出证据后才能提高验收成熟度。当前没有真实 PowerPoint 环境，工程工作仍可推进，不以此省略剩余实现；没有上传、部署或修改原方案。
