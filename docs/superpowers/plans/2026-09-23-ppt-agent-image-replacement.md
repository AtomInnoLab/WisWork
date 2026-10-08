# PPT Agent 普通图片替换

基线6bd49bd2；沿用批准的Office.js原位修改、确认、分层QA与恢复设计。隔离分支不合并/推送/部署，TDD与独立审查。

## 选择与边界

采用同页新增原生图片→校验→删除原图；页面ID保持，图片对象ID变化。旧图在候选校验及候选ID持久化前不删除。对裁剪、效果、链接、动画、组合子图片等无法保真的输入拒绝，不隐式简化。图片源使用已验证VFS PNG/JPEG（<=2MiB），不新增依赖。Office1.10。无原子CAS，不自动回滚未知写入或凭新增数量猜归属。

## 接口

- PictureSnapshot: {slideId,shapeId,geometry:{left,top,width,height},rotation,name,altTextTitle,altTextDescription,zOrderPosition,shapeIds:string[],pictureFingerprint:string,mediaDigest:string}。shapeIds为层级顺序，最多100；新图保持位置/尺寸/旋转/名称/替代文本/层级。源包inspect限定单页普通顶层p:pic、PNG/JPEG内嵌且无裁剪/效果/链接/动画等；fingerprint包含原pic语义与媒体摘要，SHA256。
- BrowserPresentationImageAdapter: inspect(slideId,shapeId,signal?)=>PictureSnapshot；replace(slideId,shapeId,base64,expected:PictureSnapshot,onInserted:(newId:string)=>Promise<void>,signal?)=>{shapeId:string}。提交前重读expected一致；只按稳定ID访问。addImage后校验候选媒体SHA256等于输入、几何0.01pt且原图未变；设置元数据/层级后onInserted持久化候选ID，重新核验旧图/候选/对象顺序再delete旧图。最终旧图不存在、新图内容与属性/层级正确才成功。失败/取消不再新增，不回滚删除，未dispatch排队错误不以同context.sync冲刷；提交后的取消只做只读归因，不能继续删除旧图。
- ImageReplacementRecord: version:1, documentId,projectId,requestId,pageId,hostSlideId,oldShapeId,assetDigest(sha256),state:'pending'|'complete',newShapeId?:string。complete要求newShapeId，所有字段严格、记录<=16KiB。key = SHA256(JSON.stringify([projectId,requestId,pageId,oldShapeId]))。文档设置最多32条/128KiB，与已有保存队列串行；读/写失败明确，更新失败恢复先前本地值。首次pending保存成功才调用adapter；候选ID保存成功才删除旧图。完成后再次调用原目标返回原结果，pending阻止重试并要求检查；历史状态不冒充当前宿主状态。
- 扩展page-editing options: vfs?:InMemoryVfs, imageAdapter?:{inspect,replace}, readImageReplacement?(key):ImageReplacementRecord|undefined, writeImageReplacement?(key,record):Promise<void>。新tool replace_presentation_page_image(project_id?,page_id,shape_id,path,explanation?)，read_presentation_image_replacement(project_id?,page_id,shape_id)。共用业务页/检查点/doc/artifact绑定；替换走现有proposals与QA钩子，验证期间核对原图和VFS源字节，闭包保留已验证素材。path仅VFS路径。state记录读出来先验证字段与当前文档/页映射，不能跨文稿返回成功。

## 单元

1. backend：powerpoint-package.ts新增只读普通图片inspect（复用受限ZIP加载），新browser-presentation-image-adapter.ts与测试。验证真实PptxGenJS图片样本、非目标/复杂图片拒绝、重复媒体和关系限制、层级/内容校验、部分/取消/候选持久化失败。scopedcommit。
2. frontend：presentation-page-editing.ts与tests接工具/确认/素材验证/日志阶段；不能让error后原目标自动重插，更新记录字段与scope一致。保持旧mock兼容，可选能力隐藏。scopedcommit。
3. root：新增presentation-image-replacement-record.ts记录类型/校验、presentation-document.ts队列设置读写、runtime注入与路由、跨层恢复和失败测试、阶段报告。scopedcommit。

## 验证/回滚

完整独立审查；所有变化的定向测试、全仓test/typecheck、变更lint、插件/Shell构建。真实Office未验收必须明示。新增独立settings键不改旧数据格式；回退工具保留pending记录，不自动清理。pending自动核验和恢复留待下一阶段，当前不可自动重放。
