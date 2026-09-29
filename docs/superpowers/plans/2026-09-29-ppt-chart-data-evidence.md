# 图表原始数据关联与冻结预检

基线703694ba。依据原方案§4.3.1科研图表忠于原数据及金融口径、证据先于表达，沿用已批准完整实施范围和隔离分支。高保证交叉审查；只改源码及合成测试，不部署或传送真实资料。

## 设计与固定接口

在PresentationPlan.slides[].chartData可选声明最多128个图表。每项{elementId,categories:string[1..50],series:[{name,points:[{value,claimId,basis}]}],unit?,currency?}；series1..10，points长度等于categories；value有限±1e15，claimId必须属于本页。basis为精确union {kind:'source',sourceId,excerptOffset,excerptText} 或 {kind:'calculation'}。sourceId必须属于该claim.sourceIds；offset非负safe整数≤12000，excerptText1..80原样；计算basis必须对应calculation claim。单位/币种各≤100。独立图表ID/每chart至少一条series/点均关联主张，重复elementId拒绝。保留旧plan形状和原SHA语义；仅显式新字段改变planDigest。旧SlideIR不加字段、原native compiler不变。

共享类型/Schema写独立模块 presentation-chart-data.ts，从presentation-plan导入chart schema/type (chart module反向仅type import PresentationPlan，避免运行时cycle)。导出 PresentationChartDataBinding、PresentationChartDataCheck、checkPresentationChartData(plan,pageId,charts)、parsePresentationChartDataCheck(value,plan?)。charts输入为actual chart elements（ElementGeometry&chart type，可helper内部Pick仅{id,categories,series}）；check不parse全部plan/deck，调用者已parse/匹配；返回{version:1,pageId,scope:'frozen_declared_data',charts:[{elementId,actual?:{categories,series:[{name,values}]},findings:[{code,claimIds}]}],checks:{data:'needs_review',sourceTruth:'not_verified',host:'not_checked'}}。actual和declared并集≤256chart；每chart codes唯一，claimIds去重属于该page，可空（未关联/缺图）。非图实际elements不得进入check。checks恒定；matched无finding只表示声明数据一致，不能标passed/真实来源通过。

Findings八类：chart_data_unbound（实际chart无声明）、chart_data_missing（声明没有实际chart）、chart_data_shape_mismatch（类别/系列名称或维度不等）、chart_data_value_mismatch（逐点actual!==expected，-0/0视相同）、chart_data_source_basis_mismatch（原样excerpt.slice(offset,len)!==text、全text不符合标准十进制/科学计数数值或Number(text)!==value）、chart_data_calculation_not_reproduced（已有受限算术reproduce status非reproduced或actual!==value）、chart_data_unit_mismatch、chart_data_currency_mismatch（chart显式单位/币种与该claim.calculation或professionalContext finance同名显式字段不同）。数值不自动推断千分位/百分号/中文单位转换；拒绝数字/符号/小数/指数邻接片段和明确逗号分组片段；已知单位/币种被省略同样产生缺口；literal regex ^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$。source basis只对计划保存原文excerpt，原始附件真实性和适用范围不认证。计算沿用arithmetic_only，不能假称inputs已逐个原文认证。claims[]原保存研究context通过既有researchbinding固定。

parseCheck严格有限大小/shape/真实numbers/整数/代码/去重。提供plan时以report.actual重建check并canonical全部比较，拒伪造已匹配/去掉缺口/添加错误claim；plan无chart声明时仍报告实际unbound。旧报告无chartData可读取；有显式plan.chartData的页必须有check，旧plan允许可选check以显示实际未关联chart。

## 三单元及root

A remote_acquisition_pc：packages/pptx-engine/src/presentation-chart-data.ts、presentation-plan.ts及package.json export ./presentation-chart-data；types/Schema/planreferential检查/chartchecker/parser及独立tests presentation-chart-data.test.ts。不改content/delivery/UI。RED→GREEN覆盖source/calculation/shape/labels/allvalue/unit/currency/missing/unbound/duplicate/identity/旧plan、范围/size。通知stable API。

B addin_build_version：presentation-content-check.ts与presentation-delivery-report.ts及对应tests。content可选chartData仅在本页actualchart或plan声明存在时返回；parseContent接可选字段，原旧shapeexact仍合法。delivery.pages.chartData同条件构建，显式plan页必须check；parse强重建check；seeds按finding.claimIds产生对应code的unverifiable issues（unbound空claim只check不造假claimID）。最多每claim8追加，cap仅页有chartData增加256，其它旧cap不变；issueDigest包含完整实际chartcheck与本页绑定，旧issue不变。supported/explained不关闭chart缺口；JSON及Markdown写完整原计划绑定、actual及finding。无新编译硬拒门，已冻结badchart能在report暴露。所有读取/问题处置沿用原CAS与QA。不碰A/UI/rootfixture。RED→GREEN oldreportexact、伪造check/丢声明check拒绝、原处置stale、无claimunbound、正确数据check仍sourceTruth未认证。

C remote_acquisition_office：apps/office-addin/src/agent/presentation-delivery-report-card.tsx/tests + powerpoints planning Skill指引（先rg实际文件，报ownership，不碰rootresearch）。中文图表check默认折叠，显示目标图表/类别/实际序列数值、每个finding中文与受影响claim、数据依据明细定位原source文段/计算，不把match显示成真实性通过。切换任务/文档不展示旧check；legacy nofield隐藏，未关联chart可见；窄面板无需新依赖。计划schema引用A现有export自动可发现；Agent指引要求规划时逐点绑定/检查/保持原数据单位/币种，不增加不存在tool。

Root：实际PC→Agent跨层将冻结chartData和原资料，检查正向/篡改数值/更换计划后旧生产不漂移、报告与nativehostZIP明细保持；统一verify/commit/build/docs，独立crossreview(A审B/root，B审A/C，C审B)。复用原fixture，避免为测试重造service。三端types+enginetypes、所有改动TS static/diff、完整相关回归→sourcecommit→Officebuild版本→阶段文档。初期老plan无field保持可读，没有自动假绑定。64%不凭测试量提升，专业实机0/20。

## 实施结果

全部单元及交叉独立复审完成，源码fdf30772；完整2897/224、最终补充77/5通过，四端类型及静态通过，Office构建buildId fdf307720f04。详见 docs/product/wiswork-ppt-agent-chart-data-evidence-progress-2026-09-29.md。整体仍64%，真实专业任务0/20，未部署。
