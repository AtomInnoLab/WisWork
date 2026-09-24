# PPT Agent 方案与实现核对

日期：2026-09-23。性质：代码与方案对照，不是新一轮功能开发或实机验收。

## 1. 核对口径与结论

方案依据：[完整方案 v1.3](/home/yykai/WisWork/docs/product/wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md)。原目录与实现工作区中的该文档内容一致。

实现依据：`codex/ppt-agent-implementation`，提交 `ea678872`，目录 `/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork`。这些实现尚未合并到 `/home/yykai/WisWork` 当前的 `codex/hide-office-bridge-status` 分支，也未部署。原目录还有其他未提交工作，本次未修改。

**总判断：新建与局部修改的工程工具链已经接通多个环节，但完整方案的 P0 尚未验收。最近几轮已提前推进部分 P1 修改能力，P0 中的生产调度、可信内容验收、真实 Office 验收和任务基准仍有明显缺口。**

不能把“某一批实现与测试通过”换算成“方案阶段已完成”，也不建议给总体完成百分比。原方案的退出条件包含真实用户任务、宿主行为和产品体验，目前证据主要是工程测试。

## 2. 按插件路线图核对

| 方案阶段                | 当前状态               | 已具备                                                                                     | 距退出标准的差距                                                                                                                      |
| ----------------------- | ---------------------- | ------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------- |
| 阶段 0：基准与可观测性  | 进行中                 | 8 页中文合成 fixture、故障注入、已有诊断机制                                               | 缺 20–30 个专业标准任务、完整业务关联指标、真实完成率/耗时/人工修正基线                                                               |
| O0：文档与会话隔离      | 进行中                 | 文档身份、Save As 绑定保护、能力协商、旧连接响应隔离；已有 PC 多连接池                     | 未证明 3 个真实 PPT 文档同时独立运行；Relay 断线仍撤销当前会话，缺完整自动重连并恢复任务的闭环                                        |
| O1：项目化与 ACP 时间线 | 进行中                 | 持久化计划、版本冲突检查、编译请求/成果恢复；项目、导入和 QA 卡片                          | 不是完整 PresentationProject/AgentRun/SlideTask 状态机；没有 Brief→研究→计划→样式→制作→验收的持久化语义时间线和后台续跑闭环           |
| O2：附件与资产          | 进行中                 | 50 MiB 附件分块续传、PC PDF/DOCX/文本解析；PNG/JPEG 解码、转 PNG、缓存、摘要及紧凑素材引用 | 网页研究、网络图片搜索/下载/换路、许可治理、更多图片格式和配额回收未接通；仍有附件及素材数量上限                                      |
| O3：SlideIR 与原生写入  | 进行中                 | 结构化 IR→PptxGenJS；原生文本/形状/图片/表格/图表；整稿编译后逐页导入、防重和检查点        | PC 仍整稿编译；没有独立逐页生成/失败页重编译；Office.js 路径尚非同一 IR 的完整对象级编译器；双端结构等价和真实可编辑性未验收          |
| O4：QA 与门禁           | 进行中                 | 几何检查、稳定页 ID 截图、真实图片进入模型、视觉结论绑定与防陈旧、修改后失效               | 内容/来源核验与保存重开 RoundTrip 未完成；当前所有已存 QA 页面都会失效，未实现受影响页精确失效；自动修复闭环不完整                    |
| O5：修改、差异与撤销    | 进行中                 | 文本替换、几何调整、普通图片替换/恢复，稳定对象定位、确认与写后核验                        | 新的稳定业务页工具依赖已生成并导入的页面映射；缺任意现稿完整 DeckBaseline/ChangeSet、选区/批量修改、持久化撤销/重应用与前后差异工作台 |
| O6：兼容发布            | 进行中                 | 能力协商与旧路径兼容、构建及协议测试基础                                                   | 未部署；缺 Windows/Mac/Desktop/Web 宿主矩阵、混合版本发布验收和部署后端到端冒烟                                                       |
| P2：品牌与规模化        | 未开始（本轮专用能力） | 可复用已有技能和 Office 高级工具                                                           | 完整品牌包治理、行业效果基准、受控并发生产、团队审阅与偏好学习未交付                                                                  |

这里的“进行中”不表示已经满足退出标准；也没有把项目既有通用工具计作新 Presentation 流程已完成。

## 3. 两条实际工作流

### 从零制作

当前可用链路：

`上传附件/图片 → PC 解析与缓存 → Agent 保存结构化计划 → Agent 提交整稿 SlideIR → PptxGenJS 编译 → 确认导入 → 按页追加与保存检查点 → Agent 逐页采集/复核 QA`

已有实现足以支持工程闭环测试，但还存在以下差别：

1. 规划字段、来源登记和引用一致性校验已实现；没有自动证明主张真实、引文准确、时效正确或计算可复现。计划中的 `reviewStatus` 只允许 `needs_review`。
2. StyleSpec 当前主要是字体及背景/文字/强调色，不是方案中的完整字号层级、网格、布局组件、图片和图表规范。
3. 逐页导入是在完整 PPTX 已编译后发生，不是“按章节边生产边验收”。导入阶段可续跑，不代表失败页能独立重编译。
4. 有计划时编译严格关联计划版本；仍兼容没有保存计划的编译请求，因此“先计划与样式契约再生产”尚不是所有入口的统一状态门禁。
5. 新任务仍依靠 Agent 调用这些工具串联；恢复项目/成果不等于恢复完整研究、生成、验收任务的执行位置。

证据：[计划契约](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/pptx-engine/src/presentation-plan.ts:13)、[Style/SlideIR](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/pptx-engine/src/presentation.ts:17)、[编译服务](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/shell/src/main/presentation-service.ts:226)、[编译回执状态](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/project-store/src/presentation-store.ts:80)。

### 修改现稿

当前新增链路：

`读取已绑定业务页/原生对象 → 提案 → 用户确认 → 核验文档及对象未变 → 修改 → 回读 → 使 QA 待复检 → 重新采集/复核`

普通图片另有 pending/candidate/complete 记录与确认恢复：有完整证据时继续删除原图或补记完成，缺证据时人工检查。

这已经增强了安全修改能力，但还不是方案第 6 章的完整现稿流程：

- 新工具只覆盖生成/导入映射已建立的页面。原有通用 Office.js/OOXML 工具仍能处理其他现稿操作，但未统一到新的完整基线、变更集与恢复契约。
- 图片原图快照是该次替换的核验凭据，不是整套 DeckBaseline；pending 恢复不是撤销。
- 文字、几何的前值检查及图片替换防重，不能替代统一 BuildReceipt、持久化保存点和“撤销最近一次 ChangeSet”。
- 替换仅支持通过严格检查的普通图片；复杂裁剪、特效、动画等保守拒绝。这个收窄合理，但不能宣称任意图片替换完成。

证据：[页绑定与修改工具](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/skills/powerpoint/presentation-page-editing.ts:381)、[图片恢复](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/skills/powerpoint/browser-presentation-image-adapter.ts:149)、[提案控制器](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/agent/proposal-controller.ts:125)。

## 4. 优先处理的偏差

### A. QA 失效范围比方案要求大

方案 §6.4、§8.9、O4 要求只使受影响页及共享样式依赖页失效。当前通用 `beforeWrite` 调用无范围的 `invalidateQa()`，将所有已保存 QA 记录中的所有页面标记为待复检。稳定文本/几何/图片修改也经过该钩子。

这不是全局禁止写入，但会造成“改一页，其他已通过页面也待复检”。建议先给已知目标页的结构化操作传入影响范围；只有无法判定范围的脚本继续使用保守全量失效。

证据：[写入钩子](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/agent/host-runtime.ts:141)、[全量失效实现](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/skills/powerpoint/presentation-document.ts:327)。

### B. 还没有按风险分级的自治策略

方案 §7.3、§13 要求普通可逆操作不频繁确认。当前新文本、几何、图片和恢复操作统一通过 proposal 确认。删除原图等动作保留确认是合理的；低风险操作尚未做到自动执行。

应先补可靠撤销与保存点，再按风险开放自动执行，不能只去掉确认按钮。

### C. 高可信定位与实际验收之间仍有缺口

方案将科研、法律、金融的证据审计列为核心，而当前主要做来源存储、主张分类和引用映射。原文定位、来源等级与冲突处理、时效/法域检查、数值口径和公式复算尚未形成验收闭环。

代码诚实保留 `sources: not_verified`、`roundTrip: not_run`；这些不应在 UI 或汇报中被编译成功、Agent 视觉通过覆盖。需要补内容/证据 QA 后，才能兑现可信内容定位。

证据：[编译检查结果](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/pptx-engine/src/presentation-compiler.ts:239)。

### D. ACP 接入不等于产品阶段时间线

已有 ACP 适配主要输出 `tool_call` / `tool_call_update`，并有项目、导入、QA 卡片。尚未形成方案列出的 `research.*`、`plan.*`、`slide.*`、`checkpoint.*` 等统一持久化语义事件及重放。不能把通用 ACP 基础记作 O1 已完成。

证据：[ACP 映射](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/agent-harness/src/acp-events.ts:25)、[项目卡](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/agent/presentation-project-card.tsx:4)、[导入进度卡](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/office-addin/src/agent/presentation-import-progress.tsx:10)。

### E. 资产链路仍有明确上限与覆盖差距

50 MiB 文件上传/PC 解析已做，并有合成 PDF 边界测试，不能继续将它列为未实现。但每文档 32 件附件、100 MiB 声明容量，IR 最多 32 个资产；图片类型、远程素材和许可能力不完整。与 O2 “没有演示文稿级图片数量硬上限”的目标不一致。

证据：[附件限制](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/shell/src/main/presentation-attachments.ts:14)、[数量/容量限制](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/shell/src/main/presentation-attachments.ts:334)、[IR 资产约束](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/pptx-engine/src/presentation.ts:123)。

## 5. 验收证据与不能据此得出的结论

上一轮完整验证截至 `ea678872`：全仓 5,951 项 Vitest 测试、类型检查、变更 Lint、Office 插件和 Shell 构建通过；另有既有脚本检查。本次为只读核对，没有重复运行全部测试。

- 5,951 是全仓测试数，不是 5,951 个 PPT 任务，也不是方案要求的完成率。
- 已有真实 PptxGenJS 生成/解包、附件解析和图片规范化证据；Office 写入、截图、取消和恢复主要使用模拟宿主。
- 8 页合成 fixture 不等于 20 个专业任务至少 16 个无需重启可交付。
- 尚无真实 PowerPoint 保存、关闭、重开后可编辑性、截图质量、3 文档并发和端到端交付率的验收记录。
- 工程恢复不保证所有不确定写入均能自动恢复。导入的不确定页、没有候选 ID/快照的图片替换仍需人工检查。

证据：[8 页合成样本](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/packages/pptx-engine/tests/fixtures/presentation-benchmark.ts:7)、[工程端到端测试](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/apps/shell/tests/presentation-end-to-end.test.ts:13)、[最近验证报告](/home/yykai/.codex/worktrees/wiswork-ppt-agent/WisWork/docs/product/wiswork-ppt-agent-image-recovery-progress-2026-09-23.md)。

## 6. 建议调整后续顺序

1. **先固定 P0 验收清单与真实任务集**：将原方案 §21 的输入、8 页输出、来源与质量条件变成可执行任务；区分工程自动化与真实 Office 人工验收。
2. **优先收窄 QA 失效范围**：这是已经可以从当前代码确认、又直接影响用户体验的偏差。
3. **补页面生产状态机与失败页重编译**：将现有计划、整稿编译、导入回执和 QA 卡片连成可恢复的生产过程；同步接入业务语义进度。
4. **补内容/证据 QA 和真实 RoundTrip**：让“通过”确实对应方案的交付要求；实机环境就绪后尽早并行验收，不必等全部功能完成。
5. **再扩展完整 P1**：任意现稿基线、ChangeSet、保存点、撤销、差异工作台和风险分级确认。网络素材、兼容格式和配额能力根据标准任务暴露的实际阻塞同步补齐。

本次仅新增核对报告，未改方案原文或实现代码。建议以后所有阶段小结统一回填本清单，避免连续局部完善掩盖 P0 退出条件仍未满足。

## 7. 本次核对后的落实记录

用户确认后，新增[P0验收清单](./wiswork-ppt-agent-p0-acceptance-2026-09-23.md)及[20项专业任务目录](./wiswork-ppt-agent-benchmark-cases-2026-09-23.md)，保持原方案为验收权威。局部QA失效偏差已在后续批次修正：仅稳定页文本/几何/图片/恢复四种已知范围操作局部失效，未知范围仍全量；详见[阶段报告](./wiswork-ppt-agent-p0-scoped-qa-progress-2026-09-23.md)。上文保留ea678872核对时的事实，不据此宣称阶段0/O4/P0已完成。

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
