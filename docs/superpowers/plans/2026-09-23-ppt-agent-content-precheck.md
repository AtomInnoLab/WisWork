# 按页内容与证据预检

基线 d69dd671。依据原方案 §11.7/O4/§21，沿用既有隔离分支；高保证、TDD、独立复审和全仓验证。原方案已获用户批准，本批补确定性预检，不宣称来源真实性、计算复现或实机内容通过。

## 设计与契约

共享纯函数读取已校验并匹配的 frozen plan/deck，只处理明确 pageId。检查 claim 文本在可见 IR 文本/表格/图表标签中的字面覆盖、来源摘录和定位缺失、引文不在已提供摘录、计算待复现。字面不匹配只提示人工核验，不能认定语义缺失。绝不联网、执行公式、解析 URI 或把引用存在当作真实核验。空白归一化，不跨元素拼成伪造匹配；备注不算可见内容。

共享模块 packages/pptx-engine/src/presentation-content-check.ts，package export ./presentation-content-check。输出 PresentationPageContentCheck：{version:1,pageId,claimIds:string[],findings:Array<{code,claimId,sourceId?}>,checks:{content:'needs_review',sources:'not_verified',calculations:'not_verified',timeliness:'not_verified',host:'not_checked'}}。code固定为 claim_text_not_found/source_excerpt_missing/source_locator_missing/quote_not_in_excerpt/calculation_not_reproduced。每 finding 含 claimId，源问题含 sourceId。32claims×3sources有界，最多256 findings；解析器严格限制字段、唯一claimIds、finding归属/重复与checks固定值。提供 checkPresentationPageContent(plan,deck,pageId) 和 parsePresentationPageContentCheck(value)。失败未知page抛not_found，非法输入沿用plan/deck校验。

PC新增只读 production_content_check，参数 operation/documentId/projectId/requestId/pageId 均必需。基于指定持久production中的 frozen plan/deck，不使用当前可变计划；生产未编译也可预检。返回 {projectId,requestId,planRevision,inputDigest,planDigest,report}，摘要直接来自记录。复用串行锁/绑定/限额/取消，不改store和状态，不触发编译或网络。

插件复用production skill，新增 check_presentation_page_content（project_id/request_id/page_id）。验证绑定、摘要、严格报告及64KiB上限，沿用epoch/document/cancel守卫；仅返回预检，不改变活动成果/QA记录/导入状态，不需确认。提供固定中文修复建议与边界说明，不把内容材料当指令。混合旧PC明确返回升级错误。

## 任务与验证

1. 共享checker+parser+测试（agent）：TDD覆盖来源缺失、字面覆盖、跨元素/备注、计算不执行、输入不变、坏报告拒绝；限定commit。
2. PC service+production路由+测试（agent）：指定任务/文档绑定、冻结计划、派生只影响目标页、只读、非法/取消及重启；限定commit。
3. 插件工具+runtime路由与测试（root）：严格响应、旧PC、取消/切文档/clear、无cache和QA副作用；真实service跨层路径。
4. 独立审查并运行全仓tests/typecheck/lint/format/licenses，再构建Addin+Shell。记录进度口径及本轮证据，提交既有分支，不合并/部署。

## 回滚与边界

仅新增只读操作与工具，无持久迁移；回退插件不影响旧记录。预检报告不是宿主写后QA、真实源核验或语义正确性判断，不新增全局写锁。下一步接证据材料核验和真实宿主验收。

## 整体进度口径

按基准阶段、O0–O6、P2九模块等权，成熟度取0/25/50/75/100；100要求原方案退出证据完整。当前25/50/50/50/75/50/50/25/0，总体约42%，是工程成熟度估算而非任务通过率。20项真实专业任务0/20。每个阶段小结沿用此口径并列依据，不随提交次数累加。
