# WisWork PPT Agent P0 验收清单

日期：2026-09-23。依据：[完整方案 v1.3](./wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md) §4–8、§10 O0–O4、§15 阶段0–4、§17、§20、§21；原方案是验收权威。本文落实验收口径，不缩减或替代原方案。

## 1. 当前结论与首批范围

**P0 进行中，未通过验收。** [20项任务目录](./wiswork-ppt-agent-benchmark-cases-2026-09-23.md)已定义；PPT-P0-01、PPT-P0-02、PPT-P0-03、PPT-P0-04、PPT-P0-05、PPT-P0-06、PPT-P0-07、PPT-P0-08、PPT-P0-09、PPT-P0-10、PPT-P0-12、PPT-P0-13、PPT-P0-14、PPT-P0-15、PPT-P0-16、PPT-P0-17、PPT-P0-18、PPT-P0-19 与 PPT-P0-20 的候选材料包已获取或自制但待领域/布局/品牌/保密边界审阅；其中 P0-06 仅为虚构合同，真实脱敏材料仍缺。仅 P0-11 完整材料待准备，真实 Office 执行均未开始。当前通过任务数未统计、完成率未测量，不能把未执行写成“0%通过率”，也不能用单元测试数替代专业任务数。

首批统一使用 PowerPoint Taskpane 发起、PC 已登录，主题加最多3份 PDF/文档、可选品牌模板；交付8页、16:9、中文为主、原生可编辑PPTX。覆盖封面、目录、文本图文、流程、对比、数据图表、总结。现稿任务使用8页基线；多文档任务每份文档独立8页。

已有合成8页样本、真实文件解析/编译测试和模拟 Office 故障注入只提供**工程证据**，不算20项专业任务通过，也不证明真实 PowerPoint 可重开、可编辑或视觉合格。

## 2. 不可缩减的退出门槛

- 固定20个任务均保留结果，至少16个无需重启任务即可交付。Taskpane按任务要求关闭/重开或连接重连不等于重建任务；记录同一project/request/checkpoint连续性。若必须重新发起制作才能交付，本项不能计入16项。
- **所有交付文件**必须在PowerPoint中保存、关闭、重开并可继续编辑；不能只打开编译器输出或解包XML。保留文件哈希及重开后的对象证据。
- 无空白页、重复页、内容越界、丢页、错误文件覆盖、不可恢复及虚假完成；不得发生不可撤销覆盖、跨文档写入或保密边界违规。P0缺陷阻止交付。
- 关键事实来源可追溯到原文、版本和位置；数字记录报告期/as-of、单位/币种及公式。来源登记和视觉pass都不代表来源核验通过。
- P1质量缺陷自动修复或明确警告，P2不应让任务失败；“核心事实无来源”的警告不能免除§21关键事实可追溯要求。
- **重做单页、撤销最近一次变更集仍属于§21首批验收**，不能以完整O5/P1后续为由删除。当前尚未实现的功能分别由任务18/19暴露；即使达到16/20，也不能把功能清单尚缺或阶段退出条件未满足表述为完整P0完成。

验收记录区分：任务结果为`未执行/通过/失败/阻塞`；材料状态为`待准备/待核验/就绪`；工程能力为`已有工程路径/部分具备/待实现`。阶段状态严格使用§20的`未开始/进行中/已完成/有条件完成/阻塞`。

## 3. 原方案—能力—证据—缺口—阶段状态

以下为本轮起始状态；局部QA失效正在同批实现，合并工程证据后仍需真实宿主验收。测试链接表示可查代码，不表示本次文档工作重新执行了测试。

| 原方案章节/阶段                     | 现能力                                                                     | 工程证据                                                                                                                                                      | 尚缺与验收任务                                                                                          | 阶段状态 |
| ----------------------------------- | -------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- | -------- |
| §15阶段0、§17：样本与指标           | 合成8页、故障注入；新增20项真实任务规格                                    | [样本](../../packages/pptx-engine/tests/fixtures/presentation-benchmark.ts)、[工程端到端](../../apps/shell/tests/presentation-end-to-end.test.ts)             | 实料、人工标注、业务诊断关联、真实耗时/成功率基线；全部20项                                             | 进行中   |
| §10 O0、§15阶段1：身份与恢复        | 文档绑定、能力协商、多连接池、回执防重                                     | [会话测试](../../apps/office-addin/tests/relay-session.test.ts)、[连接池测试](../../apps/shell/tests/office-relay-pool.test.ts)                               | 三份真实PPT同时运行、另存隔离、断线自动恢复；任务16/17/20                                               | 进行中   |
| §10 O1、§4/§7：项目工作台           | 持久化计划/成果、项目/导入/QA卡片                                          | [项目测试](../../apps/office-addin/tests/presentation-project.test.ts)、[进度测试](../../apps/office-addin/tests/presentation-import-progress.test.ts)        | 完整Project/SlideTask状态机、业务语义时间线与重放、阶段进度及控制；任务16/18/20                         | 进行中   |
| §10 O2、§15阶段2、§8.8：资料与图片  | 50MiB附件分块、PC PDF/DOCX/文本解析、PNG/JPEG规范化和引用                  | [附件集成](../../apps/shell/tests/presentation-attachments-integration.test.ts)、[素材集成](../../apps/shell/tests/presentation-assets-integration.test.ts)   | 真实50MB PDF、网络下载恢复/许可、更广格式；仍有数量限制，不满足无演示文稿级图片硬上限目标；任务10–14/16 | 进行中   |
| §10 O3、§15阶段3：IR与写入          | IR→PptxGenJS、原生对象、整稿编译后逐页导入检查点                           | [编译测试](../../packages/pptx-engine/tests/presentation-compiler.test.ts)、[逐页交付](../../apps/shell/tests/presentation-page-delivery-integration.test.ts) | 独立页面生产/重编译、完整Office.js同IR编译、双端结构等价和真实可编辑性；任务1/9/12/18                   | 进行中   |
| §10 O4、§15阶段4、§6.4/§8.9：分层QA | 几何、稳定页截图、实际模型图片传递、防陈旧复核；本轮收窄稳定操作的失效范围 | [QA测试](../../apps/office-addin/tests/presentation-qa.test.ts)、[修改后QA](../../apps/office-addin/tests/presentation-post-edit-qa.test.ts)                  | Content/证据QA、真实RoundTrip、合法重叠分类、独立失败页修复闭环；任务1–9/15/18/20                       | 进行中   |
| §21首批修改（与完整O5区分）         | 已绑定页文本/几何/普通图片替换、图片中断确认恢复                           | [稳定页工具](../../apps/office-addin/tests/presentation-page-editing.test.ts)、[图片适配](../../apps/office-addin/tests/presentation-image-adapter.test.ts)   | 任意现稿完整基线、重做单页、持久化撤销；任务13/15/18/19                                                 | 进行中   |

§15阶段1–4与O0–O4有交叉依赖；本表不因一个组件通过测试而将任一总体阶段标为完成。完整差距参见[方案核对报告](./wiswork-ppt-agent-plan-implementation-audit-2026-09-23.md)，其中局部QA的旧状态需结合本轮报告阅读。

## 4. 单次验收操作与证据

1. 冻结代码commit、PC/Addin版本、模型/provider配置和Office宿主版本/平台；登记材料权限及哈希。正文和诊断不得泄露未经授权材料。
2. 使用目录规定的材料、提示和故障步骤；从Taskpane启动，保存Brief、页计划、StyleSpec和原始来源映射。记录开始时间、首次真实页时间、结束时间及人工干预。
3. 保存运行/连接/文档/页/工具调用关联ID，关键动作录屏，记录回执和故障时间点；不能依赖聊天中的“成功”文字。
4. 执行内容、结构、几何、渲染、保存重开回读五层检查，再检查备注/演讲节奏。截图需标明页面和采集时间；小字不可辨不判视觉通过。历史QA不作为当前页新鲜证据。
5. 输出PPTX、来源/Claim Ledger、QA摘要、警告、保存点/撤销证据；逐项记录可核验事实、待人工判断及无法核验内容。
6. 每个产物保存关闭后重开，修改标题、表格单元格、图表数据并再保存；对该任务不含的对象标明不适用及理由。确认8页无重复/空白/越界。
7. 保留全部尝试，不用补跑成功覆盖首次失败；补跑单独编号并记录原因。材料更换须提升材料版本并重跑，不能挑选成功材料挪入固定分母。

## 5. 结果记录模板与指标

验收统计可使用 `node tools/ppt-agent-acceptance.mjs <records-directory>`。目录中每个 `.json` 文件必须是尝试记录数组；同一任务的 `attempt_no` 必须从 1 连续递增，先前失败记录保留，缺号会被拒绝。程序固定 20 项分母，只按每项最新尝试汇总；不足 20 项时完成率输出 `not_measured`，绝不将未执行等同失败。`passed` 记录需包含真实材料状态与清单、版本、身份、审阅人、PPTX SHA256、PowerPoint 保存重开及可编辑证据标记，且不能重启任务或留有 P0 缺陷。目录模式还要求每次 `passed` 尝试提供目录内的材料清单文件 `material_manifest`、`pptx_file` 与独立的 `reopen_evidence_file`，以及各自的 SHA256；程序逐文件核对摘要，拒绝缺失、越界路径和文件被修改。`rateThresholdMet` 只表示 16/20 数量门槛，不表示 P0 所有质量与功能条件均通过。文件摘要只能把记录绑定到实际字节，不能证明 PowerPoint 确实重开、内容可编辑或专业结论正确；这些仍须审阅原始录屏、文件和材料。测试中的合成记录也不计入真实任务结果。

Office 插件导出的当前文稿 ZIP 可先运行 `node tools/ppt-agent-stage-host-bundle.mjs <bundle.zip> <新暂存目录> PPT-P0-01`。Taskpane 可选“含 8 页宿主截图”导出；仅当当前文稿恰好 8 页且宿主支持截图时，逐页 PNG 会随包存入本机 PC，Manifest 和暂存草稿均标记为 `captured_unreviewed`，保留页序、宿主 Slide ID 和摘要。暂存命令核对 Manifest 中每个文件的摘要和任务身份，提取当前宿主 PPTX、冻结主张、历史 QA 和可选截图，并在 `acceptance-draft.json` 标出八页结构门禁与待补的真实 PowerPoint 保存重开、截图人工视觉复核、当前五层 QA、专业审阅和现场测量。暂存目录须是尚不存在的新目录；草稿没有 `outcome: passed`，不能直接放进验收记录根目录计数。截图与 PPTX 导出不是原子快照；历史 `quality.json` 不等于当前页面 QA，包内 `presentation.pptx` 也不证明已保存关闭重开。操作员完成原件核验和缺失证据后，才另行创建正式尝试记录。

操作员在 PowerPoint 中保存、关闭并重开后，可再次导出当前文稿 ZIP、暂存到另一新目录，再运行 `node tools/ppt-agent-compare-reopen.mjs <重开前暂存目录> <重开后暂存目录> <新报告.json>`。对照命令要求两份八页包属于相同文档、项目、任务及冻结计划，后包创建时间晚于前包且包摘要不同；逐个比较 PPTX 内原生部件的实际字节，忽略 ZIP 压缩和时间元数据，报告变化的部件。`exact_part_bytes` 是强字节一致证据，但命令无法证明 PowerPoint 确实关闭重开，也无法证明可编辑性或来源正确；变化时应人工检查差异，任何结果仍须配合现场录屏与重开编辑证据，不能自动计为 `passed`。

正式通过记录可选附 `artifacts.roundtrip_report_file` 与 `roundtrip_report_sha256`，两项须同时提供。统计器会核对报告文件、任务 ID、报告内最终 PPTX 摘要与交付 PPTX 摘要相同、前后包不同且时间有序，并拒绝报告伪称 `hostReopenVerified: true`。该报告是辅助技术证据；`reopen_evidence_file` 的现场证据及人工可编辑性核验仍为必填，不会因为部件字节相同而自动通过。

统计同时输出 `firstAttemptPassed` 和 `firstAttemptDeliveryRate`：按每项 `attempt_no: 1` 的结果计数，必须全部 20 项至少尝试一次后才给出比例；补跑成功不会改写首次结果。最终 `passed` 和 `completionRate` 仍按最新尝试计算，两组数值应并列报告。

目录模式还会读取通过记录的 PPTX，校验其 OOXML 包含演示文稿、8 个幻灯片引用及对应幻灯片 XML，拒绝把任意字节改名为 `.pptx` 或用非 8 页文件计入通过。每页至少须含一个原生形状、图片、图表框架、连接线或组合对象；重复引用同一页或无外部关系页面的完全相同对象树也会被拒绝。图片与图表对象还须有包内关系和非空目标部件；图表部件须是可解析的图表 XML，避免把资源丢失的页计入通过。图片等通过页面关系引用的资产可能让相同对象树呈现不同内容，因此这类页面不凭对象树判重；视觉空白、媒体真实解码/显示、关系资产内容和近似重复仍由真实 PowerPoint 截图人工核验。结构检查不替代 PowerPoint 中保存、关闭、重开和编辑的人工证据核验。

每次 `passed` 还须绑定独立的 Claim Ledger、五层 QA 报告，以及按 `page_no: 1..8` 排列的 8 张逐页 PNG 截图。字段分别为 `claim_ledger_file`/`claim_ledger_sha256`、`qa_report_file`/`qa_report_sha256`、`page_screenshots: [{page_no,file,sha256}]`。统计器核对文件路径、摘要、截图可解码性和文件不重复；Ledger、QA 和截图的专业内容及截图是否来自真实宿主仍由审阅人核验。证据 JSON 应放入子目录，目录根层的 `.json` 文件只用于尝试记录。

统计输入示例：`records/01.json` 内容为数组，失败/阻塞记录也用相同的 `case_id`、`attempt_id`、`attempt_no`、`outcome` 字段，`outcome` 取 `passed`、`failed` 或 `blocked`。通过记录还须填写以下字段；示例值只说明格式，不是验收证据：

```json
[
  {
    "case_id": "PPT-P0-01",
    "attempt_id": "2026-09-25-a",
    "attempt_no": 1,
    "outcome": "passed",
    "material_status": "ready",
    "material_manifest": "files/PPT-P0-01-material-manifest.md",
    "material_manifest_sha256": "填写材料清单文件的 64 位小写 SHA256",
    "commit_and_versions": "代码、PC、插件、PowerPoint 版本记录位置",
    "identity": "project/request/document/session 记录位置",
    "reviewer_and_date": "审阅人及日期",
    "restart_required": false,
    "p0_defects": 0,
    "measurements": {
      "started_at": "2026-09-25T00:00:00.000Z",
      "first_real_page_at": "2026-09-25T00:03:00.000Z",
      "finished_at": "2026-09-25T00:20:00.000Z",
      "manual_correction_pages": 0,
      "duplicate_writes": 0,
      "screenshot_failures": 0,
      "image_failures": 0,
      "confidentiality_violations": 0,
      "cross_document_writes": 0,
      "ratios": {
        "native_editable_objects": { "numerator": 42, "denominator": 44 },
        "critical_claims_traced": { "numerator": 7, "denominator": 8 }
      }
    },
    "artifacts": {
      "pptx_file": "files/PPT-P0-01-final.pptx",
      "pptx_sha256": "填写真实文件的 64 位小写 SHA256",
      "reopen_evidence_file": "files/PPT-P0-01-reopen.mp4",
      "reopen_evidence_sha256": "填写真实录屏的 64 位小写 SHA256",
      "claim_ledger_file": "files/PPT-P0-01-claim-ledger.json",
      "claim_ledger_sha256": "填写真实文件的 64 位小写 SHA256",
      "qa_report_file": "files/PPT-P0-01-qa.json",
      "qa_report_sha256": "填写真实文件的 64 位小写 SHA256",
      "page_screenshots": [
        {
          "page_no": 1,
          "file": "files/PPT-P0-01-page-1.png",
          "sha256": "填写第 1 页 PNG 的 SHA256"
        },
        {
          "page_no": 2,
          "file": "files/PPT-P0-01-page-2.png",
          "sha256": "填写第 2 页 PNG 的 SHA256"
        },
        {
          "page_no": 3,
          "file": "files/PPT-P0-01-page-3.png",
          "sha256": "填写第 3 页 PNG 的 SHA256"
        },
        {
          "page_no": 4,
          "file": "files/PPT-P0-01-page-4.png",
          "sha256": "填写第 4 页 PNG 的 SHA256"
        },
        {
          "page_no": 5,
          "file": "files/PPT-P0-01-page-5.png",
          "sha256": "填写第 5 页 PNG 的 SHA256"
        },
        {
          "page_no": 6,
          "file": "files/PPT-P0-01-page-6.png",
          "sha256": "填写第 6 页 PNG 的 SHA256"
        },
        {
          "page_no": 7,
          "file": "files/PPT-P0-01-page-7.png",
          "sha256": "填写第 7 页 PNG 的 SHA256"
        },
        {
          "page_no": 8,
          "file": "files/PPT-P0-01-page-8.png",
          "sha256": "填写第 8 页 PNG 的 SHA256"
        }
      ],
      "powerpoint_reopened": true,
      "editable_after_reopen": true
    }
  }
]
```

`measurements` 为可选的现场测量记录，但 `passed` 必须明确记录 `confidentiality_violations: 0` 与 `cross_document_writes: 0`；缺失不等于零。时间使用 UTC ISO 毫秒格式，计数为不超过 1,000,000 的非负整数；首次真实页面和完成时间不得早于开始时间。缺失字段代表未测量，**不能用 0 代替缺失**。统计工具对每项只使用最新一次尝试，分别报告每种测量的样本覆盖数；仅当固定 20 项均已执行、且该测量均已填写时，才输出总计、完成耗时或首张真实页耗时 P95，否则输出 `not_measured`。这些数值是记录汇总，不验证现场录屏或诊断的真实性；需将原始时间线和故障回执一并归档。

§17 质量指标可写入 `measurements.ratios`，每项填 `{ "numerator": 已观察到的符合/发生数, "denominator": 已检查的机会数 }`。允许 `0/0` 表示该任务确实没有适用对象；缺项表示没有测量。汇总仅在 20 项最新尝试全部填写该指标且总分母大于零时给出比例，同时保留原始分子、分母与覆盖数。可用键如下：

| 键                                | 分子 / 分母                                       |
| --------------------------------- | ------------------------------------------------- |
| `first_two_pages_style_revisions` | 前两页触发样式修订的任务 / 已检查任务             |
| `first_round_visual_passes`       | 首轮视觉通过页 / 已审查页                         |
| `native_editable_objects`         | 原生可编辑目标对象 / 已检查目标对象               |
| `critical_facts_sourced`          | 有来源的关键事实 / 已检查关键事实                 |
| `critical_claims_traced`          | 可定位原文、版本和页面的关键主张 / 已检查关键主张 |
| `citations_accurate`              | 与原文、页码及归因一致的引用 / 已检查引用         |
| `unsupported_factual_claims`      | 未标注且缺充分证据的事实主张 / 已检查事实主张     |
| `timely_numeric_claims`           | 满足 as-of 与报告期的数值 / 已检查数值            |
| `reproducible_calculations`       | 公式、输入、单位和币种可复算的计算 / 已检查计算   |
| `successful_recoveries`           | 中断后成功恢复并完成的任务 / 已检查中断任务       |
| `prepared_images`                 | 写页前可用的图片 / 已计划使用的图片               |
| `user_interruptions`              | 因理解或信任问题停止的任务 / 已检查任务           |
| `taskpane_recoveries`             | 重开或断线后恢复原任务 / 已检查恢复尝试           |
| `pairing_first_try`               | 正确验证码首次配对成功 / 已检查配对尝试           |
| `manual_changes_preserved`        | 续跑后保留的用户修改 / 已检查用户修改             |

其中样式修订、无依据主张和用户中断率以越低越好；其余以越高越好。填写比例前应保留审阅清单和诊断证据，统计工具不替代事实核验。

每项任务单独创建记录，未执行字段填“待执行”，不能填推测值：

```yaml
case_id: PPT-P0-01
attempt_id: 待执行
material_manifest: 待准备 # 文件名、SHA256、来源/版本/页码、授权范围、人工期望清单
commit_and_versions: 待执行 # commit、PC、Addin、Office平台/版本、模型配置
identity: 待执行 # project/request/document/session与目标slide IDs
started_at: 待执行
first_real_page_at: 待执行
finished_at: 待执行
restart_required: 待执行
manual_correction_pages: 待执行
fault_and_recovery: 待执行 # 注入点、现象、重试次数、回执/诊断关联
artifacts: 待执行 # PPTX及哈希、录屏、逐页截图、ledger、五层QA、重开编辑证据
outcome: 未执行
defects: 待执行 # P0/P1/P2、影响页、修复状态；无缺陷也需人工签署
reviewer_and_date: 待执行
```

按§17报告首次完整交付数/20、插件端端到端完成率、恢复成功率、重复写入率、原生可编辑对象比例、关键主张可追溯率、引用准确率、无依据主张率、时效合规率、计算可复现率、人工修正页数及完成耗时。记录各指标分子分母；小样本P95仅作观察并附原始耗时。跨文档/保密违规单独报告，目标为0。当前均未测量。

## 6. 阶段汇报与部署矩阵

本机故障复盘可运行 `node tools/ppt-agent-diagnostic-report.mjs <office-diagnostics.json> [relay.jsonl]`。第一份文件来自插件的诊断复制/导出；第二份可选，是 Relay 标准错误输出中逐行 JSON 的 `office_diagnostic` 日志。报告只输出失败事件的固定字段、Project/Session/Document/Page/Tool Call ID、阶段和错误码，并用会话、trace 与事件 ID 三者核对 Relay 是否观测到同一事件。`relay_observed: false` 可能是远程采样、网络丢失或日志缺失，不能单独判为上传失败。诊断文件含本机标识，分享前须检查；该工具不读取原文或判断专业内容质量，真实验收仍需录屏、文件和人工审阅。

阶段报告按§20列目标、交付、验证证据、未完成项、风险、下一步及入口条件；不能把“进行中”自动推进为下一阶段已获验收。

| 部署层                     | 当前验收状态                           | 需保留证据                                  |
| -------------------------- | -------------------------------------- | ------------------------------------------- |
| PC应用/服务与解析编译      | 工程证据已有；发布组合待选定           | 版本、包哈希、服务/模型配置、诊断           |
| Relay/协议                 | 工程证据已有；真实网络/混合版本待执行  | 版本协商、断线时间线、路由隔离记录          |
| Taskpane/Addin             | 工程证据已有；宿主加载及发布验收待执行 | manifest/版本、加载入口、能力声明           |
| PowerPoint Windows/Mac/Web | 本目录均未执行；具体版本待登记         | 宿主版本、受支持/不支持能力、截图与重开文件 |

优先补材料与执行基线，同时推进原方案的页面生产/失败页重编译、业务进度、内容证据QA和RoundTrip；保留重做单页与撤销首批缺口。完成这些证据之前不宣布P0或O0–O4验收通过。

### 页级生产后续进展（2026-09-23）

已补 PC 持久化逐页编译、失败页续跑、插件页状态和单页产物下载，详见[页级生产阶段报告](./wiswork-ppt-agent-page-production-progress-2026-09-23.md)。本批不包含新页产物的宿主写入/回读、完整业务事件流或跨任务受影响页重编译；O1/O3和P0仍为进行中。真实20项任务未执行，单页重做、撤销、内容证据QA和RoundTrip缺口继续保留。

### 页生产导入后续进展（2026-09-23）

新增完整页产物准备、独立生产导入工具、v2业务页回执和ID回读，支持已确认前缀后的续写；详见[生产导入阶段报告](./wiswork-ppt-agent-production-import-progress-2026-09-23.md)。该链路尚无逐页QA/编辑绑定、边编译边写入、实机RoundTrip或撤销，O3/P0仍为进行中。回执v2需新Taskpane，保留记录，不能用删除记录的方式降级。

### 生产页QA与编辑后续进展（2026-09-23）

生产v2业务页回执已接入截图/结构QA、Agent视觉复核、稳定文字/几何/图片编辑和图片恢复，详见[生产页QA与编辑阶段报告](./wiswork-ppt-agent-production-qa-editing-progress-2026-09-23.md)。仅已确认宿主页可操作，旧整稿和生产记录隔离；内容证据QA、真实RoundTrip、单页重做、撤销和20项实机任务仍未完成，O3/O4及P0继续为进行中。

### 几何保存点与撤销进展（2026-09-23）

新增最近一次稳定页位置/尺寸修改的持久化保存点和撤销，详见[几何撤销阶段报告](./wiswork-ppt-agent-geometry-undo-progress-2026-09-23.md)。不覆盖文字富文本、图片、整页或通用ChangeSet；不确定记录禁止自动重放。完整撤销、单页重做、真实RoundTrip和20项实机验收仍为首批缺口。

### 几何保存点恢复进展（2026-09-23）

已补pending/undo_pending几何记录的只读检查与确认恢复，目标已达到时仅补回执、原值明确匹配时才写入，模糊或混合状态需人工核对。详见[几何恢复阶段报告](./wiswork-ppt-agent-geometry-recovery-progress-2026-09-23.md)。完整撤销、单页重做、证据QA及实机RoundTrip仍未完成。

### 单页派生重编译进展（2026-09-23）

新增冻结父任务的单页修订：仅目标页重新编译，其他成功页原样复用，保留父成果；详见[单页重编译阶段报告](./wiswork-ppt-agent-page-rebuild-progress-2026-09-23.md)。修订任务禁止整批导入，宿主页替换、原页备份与整页撤销仍待实现，不能将本批视为首批单页重做已完成。

### 原页持久备份进展（2026-09-23）

单页修订已补当前宿主页原始PPTX的PC持久备份、同字节断点续传与重启后下载，严格绑定父子任务、文档和稳定页ID；详见[原页备份阶段报告](./wiswork-ppt-agent-page-backup-progress-2026-09-23.md)。备份是历史保存点，不代表宿主替换或整页撤销已完成；重复导出字节稳定性与真实Office保存重开仍待实机验证。O3/O5及P0继续进行中。

### 单页替换暂存与撤回进展（2026-09-23）

新增确认后的修订页暂存、持久pending/inserted/staged回执、已知新页的恢复核对与暂存撤回；原页及业务映射保留。详见[替换暂存阶段报告](./wiswork-ppt-agent-page-replacement-stage-progress-2026-09-23.md)。正式替换提交、整页撤销、写后验收和真实Office兼容仍未完成，O3/O5及P0保持进行中。

### 单页替换正式提交与整页撤销进展（2026-09-23）

已接通确认式正式替换、原页备份恢复、原子事务/业务映射切换和终态保存失败后的只补回执恢复；详见[正式替换与撤销阶段报告](./wiswork-ppt-agent-page-replacement-commit-progress-2026-09-23.md)。派生 prepare 仅开放给具有完整已提交映射的任务。真实 Office 兼容、写后质量验收、通用撤销和 20 项专业任务仍未完成，O3/O5 及 P0 保持进行中。

### 单页替换后的 QA 衔接（2026-09-23）

QA 工具已与有效成果状态对齐，阻止替换未决或已取代成果的旧页访问；跨层验证提交/撤销后的正确页截图、强制重新采集与非目标记录保留。详见[QA 衔接阶段报告](./wiswork-ppt-agent-replacement-qa-progress-2026-09-23.md)。截图及 Office 上下文仍为模拟，内容/证据 QA、共享样式依赖分析和真实 RoundTrip 尚未完成，P0 保持进行中。

### 内容与证据预检及统一进度（2026-09-23）

新增指定冻结生产页的只读内容/证据预检，报告主张字面覆盖、来源完整性与计算待复核问题，未自动核验真实性或宿主页。详见[阶段报告](./wiswork-ppt-agent-content-precheck-progress-2026-09-23.md)。按用户最新要求新增[进度台账](./wiswork-ppt-agent-overall-progress-2026-09-23.md)：当前整体约 **42%**（九模块等权工程成熟度估算），真实专业任务 **0/20**。后续每次小结沿用口径；早期不建议给百分比的意见保留为历史记录。

### 主张附件证据读取（2026-09-24）

已接通冻结页面→主张→来源→同文档上传附件原文窗口的只读追溯，返回精确文本偏移与摘录匹配；不将匹配视为事实/权威性/时效核验。详见[阶段报告](./wiswork-ppt-agent-claim-evidence-progress-2026-09-24.md)。整体成熟度估算仍 **42%**，真实专业任务 **0/20**；原方案退出标准不变。

### 主张证据复核记录（2026-09-24）

已接通基于当前会话已读证据的 Agent 判断持久化：支持/冲突/证据不足，PC 写前重验摘要，独立不可变记录及重启读取。详见[阶段报告](./wiswork-ppt-agent-claim-review-progress-2026-09-24.md)。整体约 **42%**，真实任务 **0/20**；按用户最新要求暂缓实机验收，继续工程和自动化测试，原退出标准未改为已通过。

### 按页多来源复核历史汇总（2026-09-24）

已接通指定冻结页面的全部主张/来源历史汇总，明确未复核、部分复核和不同判断，保留所有不可变引用；不自动认定事实或当前宿主通过。详见[阶段报告](./wiswork-ppt-agent-page-reviews-progress-2026-09-24.md)。整体仍 **42%**，真实任务 **0/20，按要求暂缓实机验收**。

### 2026-09-24：主张与来源时点元数据预检

整体 **42% → 42%**，真实专业任务 **0/20，按用户要求暂缓实机验收**。在既有冻结页面预检中报告来源时点缺失或标签不同；差异不等于过期，相同不等于时效通过，`timeliness` 仍为 `not_verified`。无新增存储或工具。见[本轮报告](./wiswork-ppt-agent-as-of-precheck-progress-2026-09-24.md)。

### 2026-09-24：可恢复生产工作台批次

按原方案O1接通后台页生产、持久任务/事件回放、暂停继续取消、跨面板恢复、多任务切换、单页保存及完整成果导入准备。真实8页工程链覆盖部分成功、失败页重试与新旧请求共存；未做Office实机验收。整体 **42% → 44%**（O1 50%→75%，其余不变），真实任务 **0/20，按要求暂缓**。完整研究/规划/宿主语义生命周期和AgentRun恢复仍待实现，不宣称O1/P0退出。见[阶段报告](./wiswork-ppt-agent-recoverable-workbench-progress-2026-09-24.md)。

### 2026-09-24：内容证据交付批次

按原方案 §4.3、阶段 F/G 接通结构化算术复现、持久问题处置、逐页摘要与完整证据导出。算术一致不代表输入来源真实；已说明不关闭机器问题；来源复核保持历史 Agent 判断。共享样式依赖、来源真实性/时效语义核验和真实 RoundTrip 仍未完成，P0 保持进行中。整体 **44% → 44%**，真实专业任务 **0/20，按要求暂缓实机验收**。详见[本轮报告](./wiswork-ppt-agent-evidence-delivery-progress-2026-09-24.md)。

### 2026-09-24：共享样式依赖与局部重审

原生母版/版式操作已根据宿主依赖快照精确作用到相关页面QA，提案确认及持久失效后重新校验依赖；读取不明/超限仍全量失效。工作台可准备指定页面重审，截图观察、确认和QA门禁保持。O4主要工程链路接通，工程成熟度 **50%→75%**，整体 **44%→47%**；真实专业任务 **0/20，暂缓实机验收**。来源语义/时效、真实RoundTrip和宿主兼容尚未验收，P0保持进行中。见[本轮报告](./wiswork-ppt-agent-shared-style-qa-progress-2026-09-24.md)。

### 2026-09-24：修改差异与撤销工作台

按原方案§6/O5补齐文字持久保存点、确认式撤销和中断恢复；四类已有记录进入工作台，动作仍经实时检查与提案确认。尚未形成任意现稿基线、图片备份撤销或统一跨类型多级历史，O5保持50%，整体 **47%→47%**。真实专业任务 **0/20，按要求暂缓实机验收**，P0仍进行中。见[阶段报告](./wiswork-ppt-agent-change-workbench-progress-2026-09-24.md)。

### 2026-09-24：图片原图备份与确认撤销

已接通原始PNG/JPEG持久备份、图片保存点、确认撤销、反向恢复与工作台入口；手工冲突/缺失备份拒绝写入，回执失败不重复插图。覆盖稳定导入映射的受支持普通原生图片，不扩张为任意现稿或多级历史。整体47%→47%、O5仍50%，真实任务0/20，实机验收继续暂缓。[阶段记录](./wiswork-ppt-agent-image-undo-progress-2026-09-24.md)。

### 2026-09-24：统一变更历史与逐项撤销

四类保存点统一持久留档；文字/几何按change_id精确读取/撤销/恢复，工作台和Agent获得单操作ChangeSet摘要，旧记录标记顺序未知。跨层覆盖重开后跨类型逐条撤销及较早记录回执恢复。仍非完整多操作ChangeSet/任意现稿基线，整体47%→47%、O5仍50%、真实任务0/20，实机验收继续暂缓。[阶段记录](./wiswork-ppt-agent-change-history-progress-2026-09-24.md)。

### 2026-09-24：现稿基线与选区上下文

当前页/选区/整套现稿读取、已读取字段的变化检测、按ID截图接入本地PowerPoint工具，无需生成成果或PC在线。只读会话基线仍非现稿持久修改/撤销/QA闭环，备注/来源/复杂对象内部未完整读取。整体 **47%→47%**，O5仍50%，真实任务 **0/20，实机验收继续暂缓**。下一步接通现稿文字/几何修改的持久保存点和撤销。[阶段记录](./wiswork-ppt-agent-deck-baseline-progress-2026-09-24.md)。

### 2026-09-24：现稿原生修改与持久撤销闭环

整体 **47%→50%（450/9）**，O5工程成熟度 **50%→75%**。现稿当前页/选区文字与几何修改已接通基线重校验、确认、写前保存点、宿主回读、重开撤销/未决恢复和局部截图历史复核；离线工作台按原生记录身份操作。全仓6642项Vitest、11项Node、53项Rust通过，独立审查无重要问题。75%表示主要工程链接通，不表示完整需求或实机验收通过；多操作ChangeSet、现稿图片/整页接线、完整基线/富文本恢复仍缺。真实任务 **0/20，按要求暂缓实机验收**。下一步多对象/多操作规划与逐项恢复。[阶段记录](./wiswork-ppt-agent-existing-edits-progress-2026-09-24.md)。
