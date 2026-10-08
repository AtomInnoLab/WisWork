# O3/O4：生产页QA和稳定编辑绑定

基线c1c5a5dc。原方案§5F、§6.4、§8.9、O3/O4、§11.8、§20/21为权威。保持隔离工作区，不合并部署。

## 目标与边界

为已导入生产页接入现有截图/结构QA、Agent视觉评审、稳定文字/几何/图片编辑及图片恢复。v2回执的业务pageIds按索引定位宿主页，不能find重复sourceSlideId。原整稿v1行为和摘要保持兼容。尚不实现内容证据审核、保存重开、单页重做、撤销或实时事件。

## 接口和交付

root在presentation-page-delivery.ts导出presentationArtifactContent(artifact):string：v1返回原pptxBase64；v2返回与现有v2digest完全一致JSON序列化(doc/project/request/revision/pages/pagePptxBase64)，严格validPages；presentationPageMapping(artifact,record,pageId):{sourceSlideId,slideId}|undefined：调用summarize验证v1/v2、身份、pageIds序列，再通过业务pageId索引取completed。调用方需检查digest并在await边界检测content变化。现有presentationImportKey统一命名空间。

单元1 backend：presentation-qa.ts、presentation-document.ts及对应tests。QA记录version1新增可选source:'production'，production key独立三段；旧记录没有source保持原key，严格拒混用。使用共享content计算摘要/防迟到，使用mapping绑定目标页。document QA读写和局部失效保持两类记录隔离与跨宿主页失效。document图片key读取/保存同步frontend可选source。TDD重复256#第二页、混同request、摘要/顺序变更、不确定页拒、局部QA、旧记录兼容，范围提交。

单元2 frontend：presentation-page-editing.ts、presentation-image-replacement-record.ts及对应tests。共享content/key/mapping用于稳定文字几何图片路径，v2重复sourceID按业务page定位；图片恢复记录新增可选source:'production'，imageReplacementKey额外可选source参数，仅生产时JSON摘要加namespace，旧算法不变；validate严格source，readRecord必须一致。与backend同步document图片保存key。TDD第二页编辑/图片恢复命名空间、错误摘要/混合版本拒绝、await变更与旧行为，范围提交。

root：sharedhelpers+host-runtime activeArtifact provider（生产选择用production.artifact，不回退旧生成），QA卡使用命名空间；移除过期未接入文案，工具提示身份来源。跨层测试证明生产→导入→QA定位/稳定编辑，独立复审和完整验证；阶段报告回填验收缺口。

## 验证/兼容

先RED后GREEN；独立审查；全仓npm test/typecheck、变更ESLint/diffcheck，测试结束后Addin/Shell构建。真实Office仍未执行。无PC/Relay升级；Office QA及图片记录新增可选字段，旧客户端可能拒绝包含新记录文档，不删回执降级；保留原整稿记录与原摘要。
