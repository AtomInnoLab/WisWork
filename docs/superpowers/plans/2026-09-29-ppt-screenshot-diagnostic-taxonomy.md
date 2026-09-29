# 截图诊断分类与下一批工程核对

基线b9b87938。原方案阶段0的失败分类与阶段4/5的截图等待恢复；当前AgentSession仅isError才记tool诊断，实际waiting_screenshot为非错误返回，因此未记录本次等待，并不是已生成成功截图事件。

1. Root在真实onToolExecuted入口，仅三类实际截图相关工具的严格、非变更waiting结构记presentation_screenshot_waiting；请求页一致、hostID边界、输出小于4KiB、精确字段、retryable=true。不改变原ToolExecution/isError/模型输出，不阻塞继续。
2. A诊断共用有限tool-error白名单（既有safe Word恢复保留），补真实QA错误分类与waiting当前能力不可用的unsupported结果。Root去重现有白名单改用同一helper，未知文本仍通用错误，不允许原文变成错误码。
3. RootAgentSession真实transport callbacks/Diagnostics路径，等待/临时失败/非法相似JSON与恢复分类的先失败后通过验证。C独立审所有变更。
4. B按原方案O3/P2剩余项核对当前实际源码，交付下一批工程优先级审计，区分已接通/真实缺失/实机证据不足，不用测试数或旧文档推测完成；只新增审计文档，不改源码。
5. Root统一相关全量、Office/PC类型/静态、Office构建，本地源码与阶段文档提交，不部署。严格64%/17候选/0真实专业保持；待用户提供的08/11原件不阻断独立工程。
