# PPT Agent：页面与图表 XML 持久恢复阶段报告

日期：2026-09-29；基线 `fd4aff62`；依据原方案 §6.4 保存点及 §14.2 中断恢复。原方案不修改；源码及最终验证见下方。

## 完成内容

`edit_slide_xml` 与 `edit_slide_chart` 从内存闭包转入独立 `package_xml` 持久事务。原始包、准备包、完整页序和全部页面 SHA 证明先保存至已配对本机 PC；每个写入阶段再次读取核验保存点。文档 settings 仅保存严格校验的引用、意图、实际页面 ID 和不可变回执，复用现有 CAS 队列、统一历史和容量预留。

导入和删除拆为两个 SDK 阶段：保存插入意图 → 保留原页导入 → 保存实际新页 ID → 核完整内容与页序 → 保存删除意图 → 删除原页 → 证明最终状态。异常不会自动逆写；未知 ACK 不触发重复插入或删除。重新打开可只读查看实际状态，用户明确确认后补齐已证明回执，或从尚未开始的阶段继续。

撤销先读验原包和当前完整证明，再持久恢复插入意图、证明实际恢复页，最后删除已应用页。恢复中的未知回执同样禁止重放。已应用但仍有 pending restore 的历史被视为未解决，阻止无关修改穿过恢复事务。

PC 新增 `presentation-package-backups.v1` 六项 PowerPoint 专有操作，复用母版 blob 实现及独立固定根目录 `presentation-package-backups`，不新增另一套分块存储引擎。不与原页或母版备份能力交叉调用。单 blob 8 MiB、分块 128 KiB，每文档 4096 blob/2 GiB 明确预算；总能力数保持既有 16 项上限。

导入所有权核验比较完整解压条目字节，保留原先明确接受的背景规范化。真实同长 FNV 碰撞媒体反例证明旧摘要 matcher 会接受，而新内容证明拒绝；不能以弱摘要作为删除原页的凭据。SDK 在业务异步检查后核完整前像，在最后读取/摘要之后同步核文档与取消状态，再排入写入；Office.js 没有原子 CAS，不宣称宿主并发窗口消失。

保留 version-1 声明式 replace_xml、1–32 唯一路径、32 KiB 程序限制、图表数据/关系保护以及完整 600 页末页支持，不新增 512 页截断。缺少已配对 PC 能力或 PowerPoint 1.8 时拒绝，不退回无保存点旧写入。

工作台、统一历史、实际页面范围 QA 和 Agent 时间线已接通。截图及复核保持历史证明属性，不冒充当前宿主视觉验收。仅核对回执及放弃未写入草稿使用现有有限 local_checkpoint 路径，实际确认验证零宿主写入、零 QA 失效、零宿主修改锁读取。

## 审查修复

独立审查复现三处问题并重新验证：

- review 输入跨异步读取可被外部改为另一状态；现在整个工具调用在首次 await 前复制，PC 保存的历史复核保持原始输入。
- 元数据提案的 operation/toolName/impact 与 Runtime 守卫不一致；现在使用真实工具名和有限 local_checkpoint 宿主标识，不扩大全局绕过规则。
- applied + pending restore 被历史门禁误作已解决；现在入账和 Agent 回执均按 pending 判未解决，持久文档回归验证拒绝无关新事务。

此外，写入前生命周期钩子曾错误使本次确认上下文失效；修复为独立视觉 epoch，clear/文档切换/取消继续使执行失效。实际 Runtime → PC → 文档 settings → 原生 SDK 插入/删除确认链验证通过。

## 验证

独立审查最终 CLEAN：两个真实缺陷复现 **2/2**，六个受影响套件 **54/54**，包括完整 600 页事务；日志 `/tmp/xml-review-proof-final.log`、`/tmp/xml-engine-integration-independent-final.log`。SDK 与完整内容证明独立 **20/20**，PC 能力/存储及旧接口兼容另有独立复核，不拼接成单次回归总数。

实际 Runtime/main 定向 **99/99**，包括确认前零写入、真实 PC 保存点、真实 settings、完整 SDK 插入删除、实际页面 ID、只读重开和两项元数据零宿主钩子；日志 `/tmp/b-package-metadata-final-tests.log`。引擎完整 **26/26** 包含 600 页末页和 32 图表路径；别名窄修复后另一次 **26/26** 不重复大规模用例。Root 持久文档及原生母版绑定 **15/15**，日志 `/tmp/package-binding-pending-green.log`。

Office、Shell、Agent Harness 类型检查通过，38 个变更 TS/TSX 的 ESLint、格式及差异检查通过。Relay 完整 Rust **2 单元 + 48 集成 = 50/50**；日志 `/tmp/package-final-rust.log`。

首次完整九工作区回归 **411 文件通过、1 文件失败；4980 通过、12 失败、1 跳过，196.96 秒**。12 项全部属于 `powerpoint-package-proposal-preimage.test.ts` 旧页面/图表夹具未提供新持久工厂，导致未生成提案；迁移夹具时保留完整页序/末页 599、不相关内容/结构/身份/顺序漂移、别名修改及写前零变更断言，不恢复无 PC 保存点旧路径。日志 `/tmp/package-full-nine-final.log`。迁移后的 18/18 定向通过，母版原 6 项保留；两类正例均核完整 600 页 stage 前像、实际导入/删除、新页 ID 及最终其余 599 页顺序。四类漂移与别名保护均保留零写入断言。最终重新运行九工作区 **412 文件、4992 项通过，1 项跳过，194.76 秒**，日志 `/tmp/package-full-nine-final-green.log`。Office、Shell、Pptx Engine、Project Store、Agent Core、Agent Harness、AI Provider、Auth、Office Bridge 全部相关套件通过；不是把先前失败结果拼接为本次通过。

夹具迁移最后独立审查 CLEAN，单文件新鲜 **18/18**；日志 `/tmp/xml-preimage-migration-independent.log`。最终 Office 类型、38 TS/TSX 静态、格式及差异检查通过。

源码提交 **c259984ca0189d4c82c220756c0a5e03b391fdaa**，40 个文件、5231 行增加/482 行删除（含测试）。Office 生产构建 **10.03 秒成功**，`dist/version.json` buildId **c259984ca018** 与源码一致，保留既有超过 500 kB 分包提示；未部署。日志 `/tmp/package-final-office-build.log`。

原方案未修改，SHA256 `b107f52d5e7da02e27220559ba1a950d10f81aa0c1a9322260c00b965a891089`。

跳过项为上批已经实际通过的原生母版同一 600 页 × 32 操作压力用例（642215 ms）；本批新 XML 的 600 页单事务正常执行，不复用它来声称本轮压力用例运行。

## 进度与下一步

严格整体 **64%（575/9，较上轮 0 个百分点）**；候选 **17/20**，真实专业 **0/20**。本批交付页面/图表 XML 恢复工程；模拟宿主与合成 PC 结果不补写真实 Office 门禁证据。

继续 Unit 2/3 母版 XML：全布局身份与依赖备份、实际导入布局映射、逐依赖应用回执、母版消失后的原包重建及全依赖撤销闭合。当前 `edit_slide_master_xml` 仍为旧内存闭包；Mac 母版 XML 排除保持，不能把单页恢复当成母版恢复。之后实施 §13.1 项目保留期、统一删除与最小审计；文档级 XML 备份无项目独占归属时保留。真实宿主、专业任务及网页登录配置退出项仍单独待验。
