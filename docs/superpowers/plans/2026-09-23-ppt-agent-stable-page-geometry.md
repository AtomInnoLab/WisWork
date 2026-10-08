# PPT Agent 稳定页面布局修改

基线48d3f027，延续已批准的Office.js原位修改和QA流程。沿用隔离分支、TDD、独立审查；不合并或部署。

## 目标

按业务页ID与原生对象ID读取及调整位置/尺寸，适用于文本框、图片等现有形状。直接按宿主ID定位，不转旧页码。对象移动/缩放后仍使用已有确认、QA失效和截图复检。图片内容替换、裁剪、旋转和组合内子对象不在本轮。

## 交付单元

1. adapter及新测试：导出PresentationPageGeometry={left,top,width,height}；readPresentationPageGeometry(slideId,shapeId,signal)=>{slideId,shapeId,geometry}；editPresentationPageGeometry(slideId,shapeId,geometry,expectedGeometry,signal)=>void。新路径PowerPoint1.10。几何均有限数值且绝对值<=100000，width/height>=0允许线条；单位pt。写前四值精确比较；成功回读允许0.01pt舍入（沿用既有布局验证阈值）。取消/sync错误后仍回读，完全应用才成功，未改变/部分/第三状态分别明确错误；不自动回滚。复用现有getItem定位和readUntilConverged，不改旧declarative逻辑。RED/GREEN+scopedcommit。
2. page-editing skill及测试：新增read_presentation_page_geometry、edit_presentation_page_geometry；后者输入geometry对象（严格四字段）与explanation?，其他标识与现有工具一致。复用现有artifact/receipt/doc/epoch映射检查及proposalcontroller。预览完整前后四值并标明pt；确认/执行前精确基线比较，写后四值0.01pt校验。新方法可选以保持旧adapter mocks，未提供时隐藏/拒绝几何工具。测试输入、错对象、hook后变化、缺失、取消、精度与回读失败。RED/GREEN+scopedcommit。
3. root运行时路由与跨层回归：业务页读取几何→确认→稳定ID修改→QA失效→结构重采；拒绝与陈旧基线不写，验证布局修改不调用index接口。阶段报告及commit。

## 边界及验收

Office不存在原子CAS，最后预读与sync间仍有竞态；部分写入不覆盖回滚。未知几何不能按0猜测。数据格式不变，回退只移除新增工具；已改文稿不自动逆转。页面边界/越界属于后续QA提示，允许有意将对象放在画布外。

独立审查全diff，目标回归、全仓test/typecheck、变更lint、Office插件与Shell构建。真实Office验证仍需实机；保留分支不推送。

## 审查修正

setter排队阶段异常，或排队后在提交前取消，直接以office_state_uncertain结束，不在同context回读sync以免主动发送半批队列。只有已经发起提交sync才进行回读归因。新增队列语义mock（赋值仅排队，sync才落地），验证这两种情况没有额外sync与宿主写入。
