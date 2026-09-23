# 首批撤销基础：稳定页几何保存点

基线e81ab5be，依据原方案§6.4、§8.7、§11.8、§14、§20/21。继续高保证/TDD/独立审查与隔离工作区，优先复用既有适配器。

## 范围

仅稳定业务页原生shape位置与尺寸的最近一次修改，可持久化保存点并在重新准备同产物后撤销。文字富文本、图片对象、整页替换与通用ChangeSet尚不覆盖，不能称完整撤销/单页重做完成。每文档保留一个几何保存点，新修改替换已结束记录；pending/undo_pending不得覆盖或自动重放。撤销需要当前几何与保存的after一致，手动变化拒绝。

## 约定接口

新presentation-geometry-change.ts定义PresentationGeometryChange={version:1,changeId:string,documentId:string,projectId:string,requestId:string,source?:'production',artifactDigest:string,pageId:string,hostSlideId:string,shapeId:string,before:{left,top,width,height},after:同,state:'pending'|'applied'|'undo_pending'|'undone'}。changeId UUID安全id，严格字段校验，geometry沿用既有界限。内容限16KiB，文档唯一记录。

DocumentBinding readGeometryChange():record|undefined; writeGeometryChange(record,expectedChange:record|undefined):Promise<void>。CAS比较完整旧记录，入队前clone输入，检查当前documentId，原子Office.settings保存/失败回滚/文档变化锁定与已有实现一致。新记录只允许undefined或applied/undone→新ID pending；同ID内容/identity不变且pending→applied→undo_pending→undone，重复同态同记录可幂等；pending/undo_pending不准覆盖。source进入immutable。

## 单元

backend负责新schema模块、document.ts、对应新tests。最小有界单记录，无通用框架。TDD字段/边界/状态/CAS/异步copy/存储失败/文档切换/旧settings无记录，scopedcommit。

frontend负责page-editing.ts与tests：options新增可选readGeometryChange/writeGeometryChange；两者配置时几何execute在宿主写前保存pending，验证后applied（无变化不写新保存点）；失败保留pending，后续几何拒绝。existing无storage调用保持旧行为。新增read_presentation_geometry_change/undo_presentation_geometry_change，仅storage+geometry adapter可用时暴露；工具输入project_id可选/page_id必需，复用当前artifact/digest/mapping身份，仅匹配记录可读/undo。read输出状态/前后几何，不推断成功。undo applied时提案，校验host after（0.01容差）与record unchanged，先保存undo_pending再写before，回读后undone；重复undone幂等无写，pending状态拒绝。新几何提案validate再次核对journal CAS。只撤销一个当前记录，不历史栈，不自动修复不确定写入。TDD前后值/重建/手动变化/错误source/保存失败/部分成功/重复undo与proposal失效。

root负责runtime/App将文档binding传入（先查所有构建options），路由与QA局部失效白名单新增undo工具；跨层验证修改→保存点→撤销→局部QA/重复防写，阶段报告和清单。独立最终审查、全仓npm test/typecheck、变更lint、测试后Addin/Shell构建。无PC/Relay/manifest变化，新Office独立settings键，旧客户端忽略但旧客户端仍可编辑使保存点失效；保留数据，不宣称全格式恢复。
