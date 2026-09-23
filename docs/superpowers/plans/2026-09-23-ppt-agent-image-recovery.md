# 图片替换中断核验与确认恢复

基线5d4b8f0a。继续已批准图片替换方案；沿用隔离分支，不合并/推送。涉及原生删除与持久化，使用高保障实现和独立审查。

## 接口与边界

ImageReplacementRecord新增可选baseline: PictureSnapshot。严格验证、绑定hostSlideId/oldShapeId、候选不在baseline.shapeIds内、容量预留保留；baseline不可变。旧记录缺baseline照常读取，只能人工检查。新替换首次pending保存baseline，不保留base64。

BrowserPresentationImageAdapter新增：
- inspectRecovery(record:ImageReplacementRecord, signal?):Promise<ImageRecoveryStatus>
- finishRecovery(record:ImageReplacementRecord, expectedStatus:'ready_to_finish'|'already_applied', signal?):Promise<{shapeId:string}>
- ImageRecoveryStatus={status:'ready_to_finish'|'already_applied'|'manual_review',reason?:string}。
只读核验先明确读取稳定slideID上对象ID列表，不能把读取失败当作对象缺失。无baseline/newShapeId、记录不是pending、或任何内容/属性/层级不符返回manual_review；读取失败/取消抛错。ready_to_finish须原图fingerprint与baseline一致、新图mediaDigest等于assetDigest、两者placement匹配baseline、完整对象顺序等于baseline在原图之后插入candidate。already_applied须旧图明确不存在、新图内容/属性与最终层级顺序匹配。任意其他状态不新增不删除不猜测归属。
finishRecovery先重新inspectRecovery并要求匹配expectedStatus；ready时仅删除oldShapeId，禁止插图与改动候选；取消或setter/排队错误后不能flush未知写入。提交后只读复验必须为already_applied才成功；already状态仅复验，不写宿主。不能提供原子CAS保证。

新tools：inspect_presentation_image_replacement、resume_presentation_image_replacement，输入project_id?、page_id、shape_id(原图ID)，resume可explanation?。保持read工具只读历史。inspect直接返回当前核验；resume只对上述两种有效pending状态建立确认proposal；manual不能提案。共用existing doc/artifact/import mapping与record严格scope验证。恢复不依赖VFS源文件或浏览器图片decode。工具可用性需adapter对应方法；resume还需read/write记录。确认validate及execute复查记录全文未变和核验状态；finish成功后保存complete。保存失败保留pending；下一次核验能判断already_applied。沿用QA hooks，恢复不自动通过QA。

## 单元

1.backend: browser-presentation-image-adapter.ts及其测试，两个新方法与状态类型。场景覆盖bothvalid/oldabsent/missingcandidate/drift/noevidence/取消/失败；scopedcommit。
2.frontend: presentation-page-editing.ts及其测试，保存baseline、核验/恢复工具、确认新鲜度与完成保存失败；旧adapter mocks兼容；scopedcommit。
3.root: record验证、document baseline不可变、runtime路由、跨层恢复集成测试、报告；scopedcommit。

## 验证

TDD、独立复审、全部变更lint、全仓test/typecheck、测试结束后Addin与Shell构建。实际Office未验收需明示。下一阶段真实Office兼容与验收，不将模拟通过宣称实机通过。
