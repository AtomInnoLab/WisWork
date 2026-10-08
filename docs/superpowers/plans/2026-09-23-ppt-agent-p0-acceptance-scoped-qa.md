# 按原方案推进 P0：验收基线与局部 QA 失效

基线 ea678872；依据完整方案v1.3阶段0、§6.4、§8.9、O4、§17、§20、§21，以及用户明确要求按方案开发。原方案为验收权威，审计及本计划不是替代规格。

## 目标与边界

建立20项专业任务的可追踪验收目录及结果记录规则，明确工程fixture与真实任务材料/Office验收的区别；关闭“修改一页使所有QA失效”偏差。当前不扩展P1，不宣称真实任务已经执行，不改变确认策略，不实现页面调度或撤销。

架构：沿用structured proposal的可信operation/toolName和已解析hostSlideId影响范围；仅白名单内的稳定页文本/几何/图片/图片恢复操作传单页scope。持久化QA与会话live截图按同一scope失效，未知操作/脚本保持全量。写入期间仍串行阻止QA采集，结束后无关页已有会话截图可继续复核。

接口：invalidateQa(hostSlideIds?:readonly string[]):Promise<void>；qaSkill.beginMutation(hostSlideIds?:readonly string[]):void。undefined表示全量；数组1..100，非空无控制字符hostID<=256，唯一；非法显式scope拒绝，排队前copy防调用者修改。按hostSlideId跨保存的project/request QA记录匹配；无匹配不写settings。beginMutation清理目标live记录，clear仍全清；错误/取消/保存失败仍禁止宿主执行，保留现有release锁行为。

Runtime仅对edit_presentation_page_text、edit_presentation_page_geometry、replace_presentation_page_image、resume_presentation_image_replacement且operation===toolName、impact.host=powerpoint、count=1、targets恰好1个有效hostID时传scope；其他情况无scope。不能将通用脚本用户输入的preview/targets当作可信页范围。

## 单元与文件

1. backend：presentation-document.ts/presentation-qa.ts，qa-binding/qa单元测试，负责持久及会话局部失效接口。TDD，测试目标/无关页、多请求同hostID、未知全量、scope边界及排队快照、保存失败；scopedcommit。
2. frontend：仅docs/product新增P0验收清单与20项任务目录/材料需求/步骤/预期/结果模板（可使用单个markdown，无新运行时）。任务覆盖科研/法律/金融、PDF转PPT、现稿、品牌、图片/图表、多文档、断线等方案场景。每项明确真实材料及宿主验收尚未准备/执行，不能把合成fixture当真实基准。退出规则沿用16/20及全部交付重开/无P0缺陷/来源可追溯；标明首批重做单页/撤销仍未实现，不能默默移出清单。维护方案章节→实现→证据→缺口→阶段状态。scopedcommit。
3. root：host-runtime路由scope、post-edit-qa两页集成tests，复审、全gate、阶段报告及核对清单更新。scopedcommit。

## 验证与发布

独立审查、目标测试、全仓test/typecheck、变更lint、测试结束后Addin/Shell构建、diffcheck。无需schema迁移；scope为可选新增参数，旧调用全量失效。回滚恢复保守全量行为。未合并/推送/部署；真实PowerPoint环境不可用时保留验收待执行，不阻塞可独立完成的工程工作，不宣称P0通过。

下一批仍按原方案推进页面生产状态机/失败页重编译与业务语义进度，再补内容/证据QA和RoundTrip；完整P1以P0门槛为前置。
