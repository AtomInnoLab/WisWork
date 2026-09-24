# 现稿基线与选区上下文

基线 b3ce6586；沿用隔离实现分支。按已批准原方案 §6.1、O5 推进，不改原方案；真实 PowerPoint 验收继续暂缓。

## 目标与边界

无需生成/导入成果即可读取当前 PowerPoint 的页序、当前页、选中页/对象、页面对象/文字/字体及样式依赖，形成有文档身份和内容摘要的 DeckBaseline。按当前页、选中页、整套（最多20页）限制范围；超限明确拒绝，不把截断当完整。截图复用现有按宿主ID读取，单独返回，基线不等于QA通过或写入授权。备注和已有来源未有可靠原生读取时明确未读取，不猜测不存在。本批不改动现稿内容，不迁移存储，不把生成页保存点伪装成任意现稿保存点。

## 架构与接口

新 browser-presentation-baseline-adapter.ts 导出 PresentationBaselineContext、PresentationBaselinePage、PresentationBaselineAdapter 和 BrowserPresentationBaselineAdapter。context: {slideIds:string[], selectedSlideIds:string[], selectedShapeIds:string[]}，第一页选中页为当前页；空选区不退回第1页。page: {slideId:string, shapes:Array<PowerPointShape & {text?:string,font?:{name:string|null,size:number|null,color:string|null}}>, masterId?:string,layoutId?:string}。缺失字体属性表示不支持/不适用；null表示混合或未知，不代表默认。adapter.readContext(signal?), readPage(slideId,signal?)；底层只读，验证数量/ID/有限坐标/文本和字节预算；只读取可文本化对象，复杂对象保留类型和范围。

新 presentation-baseline.ts 复用该adapter，documentId和现有inspectPresentationPage/inspectSlideMasters。提供 read_presentation_baseline(scope=current|selected|deck)、check_presentation_baseline(baseline_id)、read_presentation_baseline_page(baseline_id,slide_id)。双次读取范围内内容与上下文、文档身份前后检查；稳定结果才提交会话内单一基线和SHA-256摘要。check重新读取并报告变化的页面和页序/选区，不默默替换基线；旧ID/clear/取消/并发读覆盖均拒绝。页面截图必须按记录ID绑定并复核内容，返回现有图片工具结果结构。工具说明未读取字段、不可信文稿内容、非原子宿主快照、读取不授权写入。最多500个页ID、单页100对象、文字总量/JSON有明确限制。

## A：Office.js 原生读取（子任务）

拥有新browser adapter及独立测试；不改host-runtime。测试先红后绿：任意现稿/空选区/多选/原生ID/文本字体/类型/复杂对象/不支持API/超限/取消/缺页；不调用写API。保持Office.js接口和requirement set正确，参考本地office-js类型及微软官方文档。限定提交。

## Root：基线与工具接入

拥有presentation-baseline.ts、对应tests、host-runtime.ts和runtime测试。工具作为本地扩展注册，使用现有documentId，不依赖PC在线、生成成果或导入receipt。清会话/注销销毁快照和迟到请求；只保留一份有界基线。测试先红后绿：无生成成果/离线可用、文档另存/切换、内容/页序/选区变化、clear/取消/竞争、无效输入/容量、截图ID精确绑定。补齐本批进度与边界报告。

## 验证与交付

独立完整审查，修复重要问题；全仓test/typecheck/lint/licenses/format/diff，Office Add-in和Shell构建。局部模拟测试不替代实机验收。仅本地隔离分支提交，不合并/推送/部署。无schema迁移，回退代码即可移除会话工具，不删除文档用户数据。总进度按既有评分保持客观。
