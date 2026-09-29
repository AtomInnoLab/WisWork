# 原生表格内部结构语义验证

原方案 O3/O4、阶段 3：统一 IR 契约、原生可编辑、导入后读回结构一致。保留既有 PC 新建与原生包导入路径。SlideIR 当前只有字符串矩阵、字号与外框，不发明 cellStyle 或第二套图表 API。

1. A 修改 presentation-compiler.ts 与对应测试：现有生成表格应有等宽列、等高行、无合并、既定边框。PC postflight 应拒绝只改变内部结构而保留文字/外框的包。先定向缺陷测试，再实现；确认真实序列化舍入，不套用错误单位。
2. B 修改 presentation-structure-comparison.ts 与对应测试：在现有可比原生表格中读取有界行列尺寸、每行 cell 数与合并拓扑，比较源包/宿主包。文字、外框一样但内部变化必须 mismatch；无效或超限不可认证相同。合法完整对象比较；未知复杂扩展保持明确不可认证，不自动通过。
3. Root 真实编译包到实际 QA comparison 工具的集成回归：篡改行列/合并但保持文字和外框，工具不能宣告结构相同，不改变原图表/图片/来源检查。C 独立审两边及实际入口。
4. Root 统一定向、相关回归、类型和静态。源码统一提交后 Office 构建及阶段文档，未部署。严格整体 64%、候选 17/20、真实任务 0/20，不以新字段或测试提高档位。

文件边界：A packages/pptx-engine/src/presentation-compiler.ts、packages/pptx-engine/tests/presentation-compiler.test.ts；B apps/office-addin/src/skills/powerpoint/presentation-structure-comparison.ts、apps/office-addin/tests/presentation-structure-comparison.test.ts。若实际测试名称不同先报告；不得互改、commit、全量或构建。Root 拥有 plan、integration tests、台账与报告。
