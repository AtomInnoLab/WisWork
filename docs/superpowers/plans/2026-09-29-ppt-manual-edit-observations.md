# PowerPoint 手工编辑观察与批准偏好

基线 b7966996；执行原方案 §15 阶段6/§16 P2 的用户修改学习。观察差异不证明作者身份，不修改品牌或 QA。

## 架构与约束

复用只读 BrowserPresentationBaselineAdapter、配对 PC presentation.v1、现有确认控制器及偏好目录。显式采集一个精确 slideId/shapeId 的前后文字、几何及可读取聚合字体状态；先保存 before，后续刷新/重开按 observationId 读取并完成 after。原生复杂对象内部、完整富文本与真正编辑作者未知。每次宿主读取须两次一致及文档/页面身份校验；不写 PowerPoint，不自动认定人工偏好。

PC 保存不可变前后快照及 SHA256，时间由 PC 产生；正文仅在用户明确采集的本机项目保存，不入诊断。32 条观察、每条最大64KiB、总计1MiB，满额明确失败。普通偏好保存拒绝伪造观察来源；新批准入口只从同文档项目的已完成真实记录生成偏好来源。跨项目导入保留来源，旧偏好兼容，品牌规则优先。所有 PC 输入严格 schema、文档项目作用域、原文件损坏/符号链接拒绝、原子写、相同请求幂等及不同请求冲突。批准偏好和删除观察须核对 expectedBeforeDigest/expectedAfterDigest（未完成为null）；跨项目import核对原文摘要和可选来源expectedOrigin，避免相同文本的来源发生变化。

## 可审交付单元

1. shared contract + PC 保存/完成/读取/删除及批准偏好：packages/pptx-engine/src/presentation-manual-observation.ts、presentation-preference.ts，apps/shell/src/main/presentation-manual-observations.ts、presentation-preferences.ts、presentation-service.ts；对应共享与 PC tests。先 RED missing operations，再 GREEN；重开/篡改/跨文档/幂等/旧记录兼容测试。
2. Office 只读采集与批准：新增 presentation-manual-observations.ts skill，host-runtime.ts 注册/clear；使用 begin/complete/read/delete 工具及可见确认式 save preference，不从任意漂移猜作者。文档切换、取消、宿主不稳定、对象删除/类型变化均安全失败；确认前重读 after 防止继续编辑后保存过期偏好。对应技能和实际 PC 集成测试，重开继续、批准/删除、跨项目复用及新计划读取。
3. 旧 edit_slide_text 保存点路由（独立单元）：powerpoint-skill.ts 注入 durableTextEdit，host-runtime.ts 仅具备真实 PC 和持久绑定时注册；严格旧输入及0-based页面定位转精确 baseline/native ID，复用 existingEditing 的原位文字事务、原页备份、pending/CAS、撤销和恢复。无持久能力不直接写。实际确认后重开撤销、备份/日志失败零 SDK 写、取消/换文档/富文本拒绝的 RED→GREEN 测试。execute_office_js 其它旧修改路径仍待单独核对，本批不宣称所有入口已修。
4. 独立全范围审查，处理重要问题后运行相关八工作区回归/类型/lint/构建；源代码提交后本地构建，不部署。

## 回滚与迁移

新增可选来源字段与独立本机观察文件，旧记录保持可读，不批量迁移或清理。回滚源代码不删除记录；含新可选字段的偏好不保证旧客户端支持，能力按当前增强模式配置，不宣称旧发布接受新操作。删除须可见确认且仅删除指定观察；已批准偏好保留摘要来源，单独删除偏好。没有外部上传、品牌改写、自动质量认证或真实宿主验收结论。

## 验收

实际 Office skill → PC service 路径，before持久保存→重开→模拟宿主编辑→after保存→确认偏好→重开读取→确认跨项目复制→新计划读取；未确认无偏好写入、宿主零写入、品牌/QA不变。每步缺失/冲突不能被视为成功。工程验证不能证明真实人工来源、完整富文本、跨平台或降低实际人工修改率。
